using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;

namespace SeparadorDePDFs;

internal static class Program
{
    private const long MaxPdfSizeBytes = (long)(14.5 * 1024 * 1024);
    private const string PartsFolderName = "PDFs partidos";
    private const string PreferredGhostscriptExecutablePath =
        @"C:\Users\murilo.souza\Downloads\ghostxps-10.07.1-win64\gxpswin64.exe";
    private static readonly object LogLock = new();
    private static readonly Lazy<string?> GhostscriptExecutablePath = new(FindGhostscriptExecutable);

    static void Main(string[] args)
    {
        var folderPath = GetFolderPath(args);

        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            Console.WriteLine("Pasta nao encontrada.");
            return;
        }

        var pdfFiles = Directory
            .EnumerateFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly)
            .Where(HasContent)
            .Where(file => !IsGeneratedPart(file))
            .OrderBy(file => file)
            .ToArray();

        if (pdfFiles.Length == 0)
        {
            Log("Nenhum PDF encontrado na pasta informada.");
            return;
        }

        Log($"PDFs encontrados: {pdfFiles.Length}");
        Log($"Limite por parte: {FormatSize(MaxPdfSizeBytes)}");

        foreach (var pdfFile in pdfFiles)
        {
            SplitPdfIfNeeded(pdfFile);
        }

        Log("Finalizado.");
    }

    private static string? GetFolderPath(string[] args)
    {
        if (args.Length > 0)
        {
            return string.Join(' ', args).Trim('"');
        }

        Console.Write("Informe a pasta onde estao os PDFs: ");
        return Console.ReadLine()?.Trim().Trim('"');
    }

    private static void SplitPdfIfNeeded(string pdfPath)
    {
        try
        {
            var originalSize = GetFileSize(pdfPath);

            if (originalSize <= MaxPdfSizeBytes)
            {
                Log($"[OK] {Path.GetFileName(pdfPath)} ja esta dentro do limite ({FormatSize(originalSize)}).");
                return;
            }

            int pageCount;
            using (var input = OpenWithRetry(pdfPath))
            {
                pageCount = input.PageCount;
            }

            if (pageCount == 0)
            {
                Log($"[VAZIO] {pdfPath}");
                return;
            }

            var partsFolderPath = GetPartsFolderPath(pdfPath);
            Directory.CreateDirectory(partsFolderPath);
            DeleteExistingParts(pdfPath, partsFolderPath);

            var partNumber = 1;
            var createdParts = new List<string>();
            var startPageIndex = 0;

            while (startPageIndex < pageCount)
            {
                var endPageIndex = FindLargestPageRangeEnd(pdfPath, startPageIndex, pageCount - 1);

                if (endPageIndex >= startPageIndex)
                {
                    SaveRangeOrSplitFailures(
                        pdfPath,
                        partsFolderPath,
                        startPageIndex,
                        endPageIndex,
                        ref partNumber,
                        createdParts);
                    startPageIndex = endPageIndex + 1;
                    continue;
                }

                ExtractPagesSeparately(
                    pdfPath,
                    partsFolderPath,
                    startPageIndex,
                    startPageIndex,
                    ref partNumber,
                    createdParts,
                    includePageNumberInName: true);
                startPageIndex++;
            }

            var oversizedParts = createdParts
                .Where(part => GetFileSize(part) > MaxPdfSizeBytes)
                .ToArray();

            foreach (var oversizedPart in oversizedParts)
            {
                TryCompressPdfToLimit(oversizedPart);
            }

            oversizedParts = createdParts
                .Where(part => GetFileSize(part) > MaxPdfSizeBytes)
                .ToArray();

            if (oversizedParts.Length == 0)
            {
                Log($"[DIVIDIDO] {Path.GetFileName(pdfPath)} => {createdParts.Count} parte(s) em {partsFolderPath}. Original mantido.");
                return;
            }

            Log($"[ERRO LIMITE] {Path.GetFileName(pdfPath)} gerou {oversizedParts.Length} parte(s) ainda acima de 14,5 MB.");
        }
        catch (Exception ex)
        {
            Log($"[ERRO] {pdfPath} => {ex.Message}");
        }
    }

    private static void SaveRangeOrSplitFailures(
        string pdfPath,
        string partsFolderPath,
        int startPageIndex,
        int endPageIndex,
        ref int partNumber,
        List<string> createdParts)
    {
        var partPath = GetPartPath(pdfPath, partsFolderPath, partNumber);

        try
        {
            SavePageRange(pdfPath, startPageIndex, endPageIndex, partPath);
            EnsureExtractedPdfHasContent(partPath, endPageIndex - startPageIndex + 1);
            createdParts.Add(partPath);
            partNumber++;
            return;
        }
        catch (Exception ex)
        {
            DeleteIfExists(partPath);
            Log($"[FALHA FAIXA] {Path.GetFileName(pdfPath)} paginas {startPageIndex + 1}-{endPageIndex + 1}: {ex.Message}");
        }

        if (startPageIndex == endPageIndex)
        {
            ExtractPagesSeparately(
                pdfPath,
                partsFolderPath,
                startPageIndex,
                endPageIndex,
                ref partNumber,
                createdParts,
                includePageNumberInName: true);
            return;
        }

        var midPageIndex = startPageIndex + ((endPageIndex - startPageIndex) / 2);

        SaveRangeOrSplitFailures(
            pdfPath,
            partsFolderPath,
            startPageIndex,
            midPageIndex,
            ref partNumber,
            createdParts);

        SaveRangeOrSplitFailures(
            pdfPath,
            partsFolderPath,
            midPageIndex + 1,
            endPageIndex,
            ref partNumber,
            createdParts);
    }

    private static bool TryCompressPdfToLimit(string pdfPath)
    {
        if (GetFileSize(pdfPath) <= MaxPdfSizeBytes)
        {
            return true;
        }

        var ghostscriptPath = GhostscriptExecutablePath.Value;

        if (ghostscriptPath is null)
        {
            Log($"[COMPRESS SKIP] Ghostscript nao encontrado. Arquivo: {pdfPath}");
            return false;
        }

        var profiles = new[] { "/ebook", "/screen" };

        foreach (var profile in profiles)
        {
            var tempPath = Path.Combine(
                Path.GetDirectoryName(pdfPath) ?? string.Empty,
                $"{Path.GetFileNameWithoutExtension(pdfPath)}.compressed.tmp.pdf");

            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }

            var processStartInfo = new ProcessStartInfo
            {
                FileName = ghostscriptPath,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            processStartInfo.ArgumentList.Add("-sDEVICE=pdfwrite");
            processStartInfo.ArgumentList.Add("-dCompatibilityLevel=1.4");
            processStartInfo.ArgumentList.Add($"-dPDFSETTINGS={profile}");
            processStartInfo.ArgumentList.Add("-dNOPAUSE");
            processStartInfo.ArgumentList.Add("-dQUIET");
            processStartInfo.ArgumentList.Add("-dBATCH");
            processStartInfo.ArgumentList.Add("-dDetectDuplicateImages=true");
            processStartInfo.ArgumentList.Add("-dCompressFonts=true");
            processStartInfo.ArgumentList.Add("-dSubsetFonts=true");
            processStartInfo.ArgumentList.Add($"-sOutputFile={tempPath}");
            processStartInfo.ArgumentList.Add(pdfPath);

            using var process = Process.Start(processStartInfo);

            if (process is null)
            {
                continue;
            }

            process.WaitForExit();

            if (process.ExitCode != 0 || !File.Exists(tempPath))
            {
                Log($"[COMPRESS ERRO] {pdfPath} ({profile}) => {process.StandardError.ReadToEnd()}");
                continue;
            }

            var originalSize = GetFileSize(pdfPath);
            var compressedSize = GetFileSize(tempPath);

            if (compressedSize < originalSize)
            {
                File.Delete(pdfPath);
                File.Move(tempPath, pdfPath);
                Log($"[COMPRESS] {Path.GetFileName(pdfPath)} {FormatSize(originalSize)} => {FormatSize(compressedSize)} ({profile})");
            }
            else
            {
                File.Delete(tempPath);
            }

            if (GetFileSize(pdfPath) <= MaxPdfSizeBytes)
            {
                return true;
            }
        }

        return false;
    }

    private static void ExtractPagesSeparately(
        string pdfPath,
        string partsFolderPath,
        int startPageIndex,
        int endPageIndex,
        ref int partNumber,
        List<string> createdParts,
        bool includePageNumberInName)
    {
        for (var pageIndex = startPageIndex; pageIndex <= endPageIndex; pageIndex++)
        {
            var currentPartNumber = partNumber++;
            var singlePagePath = includePageNumberInName
                ? GetPagePartPath(pdfPath, partsFolderPath, currentPartNumber, pageIndex)
                : GetPartPath(pdfPath, partsFolderPath, currentPartNumber);

            try
            {
                SavePageRange(pdfPath, pageIndex, pageIndex, singlePagePath);
                EnsureExtractedPdfHasContent(singlePagePath, 1);
            }
            catch (Exception ex)
            {
                DeleteIfExists(singlePagePath);
                Log($"[FALHA PAGINA] {Path.GetFileName(pdfPath)} pagina {pageIndex + 1} via PDFsharp: {ex.Message}");

                singlePagePath = GetPagePartPath(pdfPath, partsFolderPath, currentPartNumber, pageIndex);

                if (!TryExtractPageWithGhostscript(pdfPath, pageIndex, singlePagePath))
                {
                    Log($"[PAGINA NAO EXTRAIDA] {Path.GetFileName(pdfPath)} pagina {pageIndex + 1}");
                    continue;
                }
            }

            TryCompressPdfToLimit(singlePagePath);
            createdParts.Add(singlePagePath);

            if (GetFileSize(singlePagePath) > MaxPdfSizeBytes)
            {
                Log($"[AVISO] Pagina {pageIndex + 1} de {Path.GetFileName(pdfPath)} ficou acima do limite mesmo sozinha.");
            }
        }
    }

    private static bool TryExtractPageWithGhostscript(string sourcePdfPath, int pageIndex, string outputPath)
    {
        var ghostscriptPath = GhostscriptExecutablePath.Value;

        if (ghostscriptPath is null)
        {
            Log($"[GS SKIP] Ghostscript nao encontrado para extrair pagina {pageIndex + 1}.");
            return false;
        }

        DeleteIfExists(outputPath);

        var processStartInfo = new ProcessStartInfo
        {
            FileName = ghostscriptPath,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };

        processStartInfo.ArgumentList.Add("-sDEVICE=pdfwrite");
        processStartInfo.ArgumentList.Add("-dCompatibilityLevel=1.4");
        processStartInfo.ArgumentList.Add("-dNOPAUSE");
        processStartInfo.ArgumentList.Add("-dQUIET");
        processStartInfo.ArgumentList.Add("-dBATCH");
        processStartInfo.ArgumentList.Add($"-dFirstPage={pageIndex + 1}");
        processStartInfo.ArgumentList.Add($"-dLastPage={pageIndex + 1}");
        processStartInfo.ArgumentList.Add($"-sOutputFile={outputPath}");
        processStartInfo.ArgumentList.Add(sourcePdfPath);

        using var process = Process.Start(processStartInfo);

        if (process is null)
        {
            return false;
        }

        process.WaitForExit();

        if (process.ExitCode == 0 && File.Exists(outputPath) && GetFileSize(outputPath) > 0)
        {
            try
            {
                EnsureExtractedPdfHasContent(outputPath, 1);
                Log($"[GS PAGINA] {Path.GetFileName(sourcePdfPath)} pagina {pageIndex + 1}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"[GS VAZIO] {Path.GetFileName(sourcePdfPath)} pagina {pageIndex + 1}: {ex.Message}");
            }
        }

        Log($"[GS ERRO] {Path.GetFileName(sourcePdfPath)} pagina {pageIndex + 1}: {process.StandardError.ReadToEnd()}");
        DeleteIfExists(outputPath);
        return false;
    }

    private static string? FindGhostscriptExecutable()
    {
        var executableNames = new[] { "gswin64c.exe", "gswin32c.exe", "gs.exe" };
        var pathFolders = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        foreach (var folder in pathFolders)
        {
            foreach (var executableName in executableNames)
            {
                var candidate = Path.Combine(folder.Trim(), executableName);

                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        var programFolders = new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        };

        foreach (var programFolder in programFolders.Where(f => !string.IsNullOrWhiteSpace(f)))
        {
            var ghostscriptFolder = Path.Combine(programFolder, "gs");

            if (!Directory.Exists(ghostscriptFolder))
            {
                continue;
            }

            foreach (var candidate in Directory.EnumerateFiles(ghostscriptFolder, "gswin*c.exe", SearchOption.AllDirectories))
            {
                return candidate;
            }
        }

        if (File.Exists(PreferredGhostscriptExecutablePath))
        {
            Log($"[GS AVISO] Usando {PreferredGhostscriptExecutablePath}, mas esse executavel e GhostXPS e pode nao ler PDF de entrada.");
            return PreferredGhostscriptExecutablePath;
        }

        return null;
    }

    private static int FindLargestPageRangeEnd(string sourcePdfPath, int startPageIndex, int lastPageIndex)
    {
        var low = startPageIndex;
        var high = lastPageIndex;
        var best = startPageIndex - 1;

        while (low <= high)
        {
            var mid = low + ((high - low) / 2);
            long candidateSize;

            try
            {
                candidateSize = GetPageRangePdfSize(sourcePdfPath, startPageIndex, mid);
            }
            catch (Exception ex)
            {
                Log($"[FALHA CALCULO] {Path.GetFileName(sourcePdfPath)} paginas {startPageIndex + 1}-{mid + 1}: {ex.Message}");
                high = mid - 1;
                continue;
            }

            if (candidateSize <= MaxPdfSizeBytes)
            {
                best = mid;
                low = mid + 1;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    private static long GetPageRangePdfSize(string sourcePdfPath, int startPageIndex, int endPageIndex)
    {
        using var stream = new MemoryStream();
        SavePageRange(sourcePdfPath, startPageIndex, endPageIndex, stream);
        return stream.Length;
    }

    private static void SavePageRange(string sourcePdfPath, int startPageIndex, int endPageIndex, string outputPath)
    {
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        using var output = CreatePdfDocument();
        AddPages(sourcePdfPath, startPageIndex, endPageIndex, output);
        output.Save(outputPath);
    }

    private static void SavePageRange(string sourcePdfPath, int startPageIndex, int endPageIndex, Stream outputStream)
    {
        using var output = CreatePdfDocument();
        AddPages(sourcePdfPath, startPageIndex, endPageIndex, output);
        output.Save(outputStream, false);
    }

    private static void AddPages(string sourcePdfPath, int startPageIndex, int endPageIndex, PdfDocument output)
    {
        using var input = OpenWithRetry(sourcePdfPath);

        for (var pageIndex = startPageIndex; pageIndex <= endPageIndex; pageIndex++)
        {
            output.AddPage(input.Pages[pageIndex]);
        }
    }

    private static PdfDocument CreatePdfDocument()
    {
        var document = new PdfDocument();
        document.Options.CompressContentStreams = true;
        document.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
        document.Options.UseFlateDecoderForJpegImages = PdfUseFlateDecoderForJpegImages.Automatic;
        document.Options.EnableCcittCompressionForBilevelImages = true;
        return document;
    }

    private static PdfDocument OpenWithRetry(string file, int maxRetries = 3)
    {
        for (var attempt = 0; attempt < maxRetries; attempt++)
        {
            try
            {
                return PdfReader.Open(file, PdfDocumentOpenMode.Import);
            }
            catch
            {
                Thread.Sleep(200);
            }
        }

        throw new InvalidOperationException($"Falha ao abrir PDF: {file}");
    }

    private static void EnsureExtractedPdfHasContent(string pdfPath, int expectedPageCount)
    {
        if (!File.Exists(pdfPath) || GetFileSize(pdfPath) == 0)
        {
            throw new InvalidOperationException("PDF extraido nao foi criado ou ficou com 0 bytes.");
        }

        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);

        if (document.PageCount != expectedPageCount)
        {
            throw new InvalidOperationException($"PDF extraido ficou com {document.PageCount} pagina(s), esperado {expectedPageCount}.");
        }

        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var page = document.Pages[pageIndex];

            if (page.Contents.Elements.Count == 0)
            {
                throw new InvalidOperationException($"Pagina {pageIndex + 1} do PDF extraido ficou sem stream de conteudo.");
            }
        }
    }

    private static string GetPartsFolderPath(string pdfPath)
    {
        var directory = Path.GetDirectoryName(pdfPath) ?? string.Empty;
        return Path.Combine(directory, PartsFolderName);
    }

    private static void DeleteExistingParts(string pdfPath, string partsFolderPath)
    {
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfPath);

        foreach (var partPath in Directory.EnumerateFiles(partsFolderPath, $"{fileNameWithoutExtension} - parte *.pdf", SearchOption.TopDirectoryOnly))
        {
            File.Delete(partPath);
        }
    }

    private static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    private static string GetPartPath(string pdfPath, string partsFolderPath, int partNumber)
    {
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfPath);
        return Path.Combine(partsFolderPath, $"{fileNameWithoutExtension} - parte {partNumber}.pdf");
    }

    private static string GetPagePartPath(string pdfPath, string partsFolderPath, int partNumber, int pageIndex)
    {
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfPath);
        return Path.Combine(partsFolderPath, $"{fileNameWithoutExtension} - parte {partNumber} - pagina {pageIndex + 1}.pdf");
    }

    private static bool IsGeneratedPart(string file)
    {
        var fileName = Path.GetFileNameWithoutExtension(file);
        return fileName.Contains(" - parte ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasContent(string file)
    {
        try
        {
            return GetFileSize(file) > 0;
        }
        catch (Exception ex)
        {
            Log($"[ERRO TAMANHO] {file} => {ex.Message}");
            return false;
        }
    }

    private static long GetFileSize(string filePath)
    {
        return new FileInfo(filePath).Length;
    }

    private static string FormatSize(long bytes)
    {
        return $"{bytes / 1024d / 1024d:0.00} MB";
    }

    private static void Log(string message)
    {
        lock (LogLock)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }
    }
}
