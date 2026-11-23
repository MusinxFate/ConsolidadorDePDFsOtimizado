using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace SeparacaoDeArquivosPorCompensacao;

/// <summary>
/// Leitor mínimo de XLSX para evitar automação do Excel e dependências externas.
/// Lê valores de células da primeira aba, incluindo shared strings e inline strings.
/// </summary>
internal static class XlsxReader
{
    private static readonly XNamespace SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace OfficeRelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    public static WorksheetData ReadFirstWorksheet(string workbookPath)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        var workbookDocument = LoadRequiredXml(archive, "xl/workbook.xml");
        var firstSheet = workbookDocument.Descendants(SpreadsheetNamespace + "sheet").FirstOrDefault()
            ?? throw new InvalidDataException("A planilha não contém nenhuma aba.");

        var sheetName = (string?)firstSheet.Attribute("name") ?? "Primeira aba";
        var relationshipId = (string?)firstSheet.Attribute(OfficeRelationshipsNamespace + "id")
            ?? throw new InvalidDataException("A primeira aba não possui um relacionamento válido.");

        var relationshipsDocument = LoadRequiredXml(archive, "xl/_rels/workbook.xml.rels");
        var relationship = relationshipsDocument
            .Descendants(PackageRelationshipsNamespace + "Relationship")
            .FirstOrDefault(item => string.Equals(
                (string?)item.Attribute("Id"), relationshipId, StringComparison.Ordinal))
            ?? throw new InvalidDataException("Não foi possível localizar o arquivo XML da primeira aba.");

        var target = (string?)relationship.Attribute("Target")
            ?? throw new InvalidDataException("O relacionamento da primeira aba não possui destino.");
        var worksheetDocument = LoadRequiredXml(archive, ResolveWorkbookTarget(target));
        var rows = ReadRows(worksheetDocument, ReadSharedStrings(archive));
        return new WorksheetData(sheetName, rows);
    }

    private static IReadOnlyList<WorksheetRow> ReadRows(
        XDocument worksheetDocument,
        IReadOnlyList<string> sharedStrings)
    {
        var rows = new List<WorksheetRow>();
        var inferredRowNumber = 0;

        foreach (var rowElement in worksheetDocument.Descendants(SpreadsheetNamespace + "row"))
        {
            var rowNumber = TryParsePositiveInteger((string?)rowElement.Attribute("r")) ?? ++inferredRowNumber;
            inferredRowNumber = Math.Max(inferredRowNumber, rowNumber);
            var values = new Dictionary<int, string>();
            var inferredColumnNumber = 0;

            foreach (var cell in rowElement.Elements(SpreadsheetNamespace + "c"))
            {
                var columnNumber = GetColumnNumber((string?)cell.Attribute("r")) ?? ++inferredColumnNumber;
                inferredColumnNumber = Math.Max(inferredColumnNumber, columnNumber);
                values[columnNumber] = ReadCellValue(cell, sharedStrings);
            }

            rows.Add(new WorksheetRow(rowNumber, values));
        }

        return rows;
    }

    private static string ReadCellValue(XElement cell, IReadOnlyList<string> sharedStrings)
    {
        var cellType = (string?)cell.Attribute("t");
        if (string.Equals(cellType, "inlineStr", StringComparison.Ordinal))
            return string.Concat(cell.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value));

        var rawValue = cell.Element(SpreadsheetNamespace + "v")?.Value ?? string.Empty;
        if (string.Equals(cellType, "s", StringComparison.Ordinal) &&
            int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out var index) &&
            index >= 0 && index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }

        if (string.Equals(cellType, "b", StringComparison.Ordinal))
            return rawValue == "1" ? "VERDADEIRO" : "FALSO";
        return rawValue;
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
            return [];

        using var stream = entry.Open();
        var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        return document.Descendants(SpreadsheetNamespace + "si")
            .Select(item => string.Concat(item.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value)))
            .ToArray();
    }

    private static XDocument LoadRequiredXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path)
            ?? throw new InvalidDataException($"O arquivo Excel não contém a parte obrigatória '{path}'.");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    private static string ResolveWorkbookTarget(string target)
    {
        var normalized = target.Replace('\\', '/').TrimStart('/');
        if (normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            return normalized;

        var segments = new List<string> { "xl" };
        foreach (var segment in normalized.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count == 0)
                    throw new InvalidDataException("O caminho interno da primeira aba é inválido.");
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static int? GetColumnNumber(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
            return null;
        var columnNumber = 0;
        foreach (var character in cellReference)
        {
            if (!char.IsAsciiLetter(character))
                break;
            columnNumber = checked(columnNumber * 26 + (char.ToUpperInvariant(character) - 'A' + 1));
        }
        return columnNumber > 0 ? columnNumber : null;
    }

    private static int? TryParsePositiveInteger(string? value) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) && parsed > 0
            ? parsed
            : null;
}

internal sealed record WorksheetData(string Name, IReadOnlyList<WorksheetRow> Rows);

internal sealed record WorksheetRow(int RowNumber, IReadOnlyDictionary<int, string> Values)
{
    public string GetValue(int columnNumber) =>
        Values.TryGetValue(columnNumber, out var value) ? value : string.Empty;
}
