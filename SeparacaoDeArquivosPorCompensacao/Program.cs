using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace SeparacaoDeArquivosPorCompensacao;

internal static partial class Program
{
    private const string DefaultSourcePath =
        @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\13. PAFs\BOOKS";
    private const string DefaultSpreadsheetName = "PlanilhaPaths.xlsx";
    private const string ReportDirectory = @"C:\Users\murilo.souza\Desktop\Manu";
    private static readonly object ConsoleLock = new();

    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        if (!AppOptions.TryParse(args, out var options, out var optionsError))
        {
            Error(optionsError!);
            PrintUsage();
            return 2;
        }

        if (options.ShowHelp)
        {
            PrintUsage();
            return 0;
        }

        var spreadsheetPath = options.SpreadsheetPath ?? FindDefaultSpreadsheet();
        if (spreadsheetPath is null || !File.Exists(spreadsheetPath))
        {
            Error($"A planilha '{DefaultSpreadsheetName}' não foi encontrada. Informe o caminho com --planilha.");
            return 1;
        }

        var sourceRoot = Path.GetFullPath(options.SourcePath);
        if (!Directory.Exists(sourceRoot))
        {
            Error($"A pasta de origem não foi encontrada ou não está acessível: {sourceRoot}");
            return 1;
        }

        Console.WriteLine($"Planilha: {Path.GetFullPath(spreadsheetPath)}");
        Console.WriteLine($"Origem:   {sourceRoot}");
        Console.WriteLine("Operação: COPIAR (os arquivos de origem serão mantidos)");
        Console.WriteLine($"Modo:     {(options.DryRun ? "SIMULAÇÃO (nenhum arquivo será alterado)" : "EXECUÇÃO")}");
        Console.WriteLine($"Paralelismo: {options.Parallelism} reclamantes simultâneos");
        Console.WriteLine();

        try
        {
            var worksheet = XlsxReader.ReadFirstWorksheet(spreadsheetPath);
            Console.WriteLine($"Aba lida: {worksheet.Name}");
            var result = new ProcessingResult { Parallelism = options.Parallelism };
            var requests = ReadRequests(worksheet, result).ToArray();

            if (requests.Length == 0)
            {
                RegisterError(result,
                    "Nenhuma linha válida foi encontrada nas colunas A (Reclamante) e C (Caminho).");
                PrintSummary(result, requests.Length);
                var emptyReportPath = WriteReport(
                    spreadsheetPath,
                    worksheet.Name,
                    sourceRoot,
                    requests.Length,
                    options.DryRun,
                    result);
                Console.WriteLine($"Relatório: {emptyReportPath}");
                return 1;
            }

            using var cancellation = new CancellationTokenSource();
            Console.CancelKeyPress += (_, eventArgs) =>
            {
                eventArgs.Cancel = true;
                cancellation.Cancel();
                Warning("Cancelamento solicitado; terminando o arquivo atual.");
            };

            var stopwatch = Stopwatch.StartNew();
            ProcessRequests(requests, sourceRoot, options, cancellation.Token, result);
            stopwatch.Stop();
            result.Elapsed = stopwatch.Elapsed;
            PrintSummary(result, requests.Length);
            var reportPath = WriteReport(
                spreadsheetPath,
                worksheet.Name,
                sourceRoot,
                requests.Length,
                options.DryRun,
                result);
            Console.WriteLine($"Relatório: {reportPath}");
            return result.Errors == 0 ? 0 : 1;
        }
        catch (OperationCanceledException)
        {
            Warning("Operação cancelada.");
            return 130;
        }
        catch (Exception ex)
        {
            Error($"Falha no processamento: {ex.GetType().Name}: {ex.Message}");
            return 1;
        }
    }

    private static IEnumerable<TransferRequest> ReadRequests(
        WorksheetData worksheet,
        ProcessingResult result)
    {
        foreach (var row in worksheet.Rows.Where(row => row.RowNumber >= 2))
        {
            var claimant = CleanWhitespace(row.GetValue(1));
            var destinationBase = CleanPath(row.GetValue(3));

            if (claimant.Length == 0 && destinationBase.Length == 0)
                continue;
            if (claimant.Length == 0)
            {
                RegisterWarning(result, $"Linha {row.RowNumber}: coluna A vazia; linha ignorada.");
                continue;
            }
            if (destinationBase.Length == 0)
            {
                RegisterWarning(result, $"Linha {row.RowNumber}: coluna C vazia; linha ignorada.");
                continue;
            }
            if (!IsValidFolderName(claimant, out var reason))
            {
                RegisterWarning(result,
                    $"Linha {row.RowNumber}: reclamante inválido ({reason}); linha ignorada.");
                continue;
            }

            string fullDestinationBase;
            try
            {
                fullDestinationBase = Path.GetFullPath(destinationBase);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                RegisterWarning(result,
                    $"Linha {row.RowNumber}: caminho inválido na coluna C ({ex.Message}); linha ignorada.");
                continue;
            }

            yield return new TransferRequest(
                row.RowNumber,
                claimant,
                NormalizeForComparison(claimant),
                Path.Combine(fullDestinationBase, claimant));
        }
    }

    private static void ProcessRequests(
        IReadOnlyCollection<TransferRequest> requests,
        string sourceRoot,
        AppOptions options,
        CancellationToken cancellationToken,
        ProcessingResult result)
    {
        var sourceFolders = BuildSourceFolderIndex(sourceRoot);
        var conflictingRows = FindConflictingRows(requests);
        var processedMappings = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        Parallel.ForEach(
            requests,
            new ParallelOptions
            {
                MaxDegreeOfParallelism = options.Parallelism,
                CancellationToken = cancellationToken
            },
            request => ProcessRequest(
                request,
                sourceFolders,
                conflictingRows,
                processedMappings,
                options,
                cancellationToken,
                result));
    }

    private static void ProcessRequest(
        TransferRequest request,
        IReadOnlyDictionary<string, List<string>> sourceFolders,
        IReadOnlySet<int> conflictingRows,
        ConcurrentDictionary<string, byte> processedMappings,
        AppOptions options,
        CancellationToken cancellationToken,
        ProcessingResult result)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (conflictingRows.Contains(request.RowNumber))
        {
            RegisterError(result,
                $"Linha {request.RowNumber}: '{request.Claimant}' aparece com destinos diferentes; linha ignorada.");
            return;
        }

        var mappingKey = $"{request.ComparisonKey}\0{request.DestinationFolder}";
        if (!processedMappings.TryAdd(mappingKey, 0))
        {
            RegisterWarning(result, $"Linha {request.RowNumber}: mapeamento duplicado; linha ignorada.");
            result.IncrementDuplicateRows();
            return;
        }

        if (!sourceFolders.TryGetValue(request.ComparisonKey, out var matches))
        {
            RegisterWarning(result,
                $"Linha {request.RowNumber}: pasta de origem não encontrada para '{request.Claimant}'.");
            result.IncrementSourceFoldersNotFound();
            return;
        }
        if (matches.Count > 1)
        {
            RegisterError(result,
                $"Linha {request.RowNumber}: mais de uma pasta corresponde a '{request.Claimant}': " +
                string.Join(" | ", matches));
            return;
        }

        var sourceFolder = matches[0];
        if (IsSameOrInside(request.DestinationFolder, sourceFolder))
        {
            RegisterError(result,
                $"Linha {request.RowNumber}: o destino não pode ser a origem nem ficar dentro dela: " +
                request.DestinationFolder);
            return;
        }

        string[] files;
        try
        {
            files = Directory.EnumerateFiles(sourceFolder, "*", new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible = false,
                AttributesToSkip = FileAttributes.ReparsePoint,
                ReturnSpecialDirectories = false
            }).ToArray();
        }
        catch (Exception ex)
        {
            RegisterError(result,
                $"Linha {request.RowNumber} - {request.Claimant}: " +
                $"não foi possível enumerar os arquivos de '{sourceFolder}': {ex.Message}");
            return;
        }

        if (files.Length == 0)
            RegisterWarning(result,
                $"Linha {request.RowNumber} - {request.Claimant}: a pasta de origem está vazia: {sourceFolder}");

        if (!options.DryRun)
        {
            try
            {
                Directory.CreateDirectory(request.DestinationFolder);
            }
            catch (Exception ex)
            {
                RegisterError(result,
                    $"Linha {request.RowNumber} - {request.Claimant}: " +
                    $"não foi possível criar '{request.DestinationFolder}': {ex.Message}");
                return;
            }
        }

        result.IncrementMatchedFolders();
        var processedInRequest = 0;
        foreach (var sourceFile in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relativePath = Path.GetRelativePath(sourceFolder, sourceFile);
            var destinationFile = Path.Combine(request.DestinationFolder, relativePath);

            if (File.Exists(destinationFile))
            {
                RegisterWarning(result,
                    $"Linha {request.RowNumber} - já existe e não foi sobrescrito: {destinationFile}");
                result.IncrementCollisions();
                continue;
            }

            if (options.DryRun)
            {
                result.AddCopiedFile($"{sourceFile} -> {destinationFile}");
                processedInRequest++;
                continue;
            }

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(destinationFile)!);
                File.Copy(sourceFile, destinationFile, overwrite: false);
                result.AddCopiedFile($"{sourceFile} -> {destinationFile}");
                processedInRequest++;
            }
            catch (Exception ex)
            {
                RegisterError(result,
                    $"Linha {request.RowNumber} - falha ao copiar '{sourceFile}' para " +
                    $"'{destinationFile}': {ex.GetType().Name}: {ex.Message}");
            }
        }

        Info(
            $"[Linha {request.RowNumber}] {request.Claimant}: " +
            $"{processedInRequest:N0} de {files.Length:N0} arquivo(s) " +
            $"{(options.DryRun ? "planejado(s)" : "copiado(s)")}.");
    }

    private static Dictionary<string, List<string>> BuildSourceFolderIndex(string sourceRoot)
    {
        var index = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var folder in Directory.EnumerateDirectories(sourceRoot, "*", SearchOption.TopDirectoryOnly))
        {
            var key = NormalizeForComparison(Path.GetFileName(folder));
            if (!index.TryGetValue(key, out var matches))
            {
                matches = [];
                index.Add(key, matches);
            }
            matches.Add(folder);
        }
        return index;
    }

    private static HashSet<int> FindConflictingRows(IEnumerable<TransferRequest> requests)
    {
        var conflicts = new HashSet<int>();
        foreach (var group in requests.GroupBy(request => request.ComparisonKey))
        {
            var destinations = group.Select(request => Path.TrimEndingDirectorySeparator(request.DestinationFolder))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (destinations.Length <= 1)
                continue;
            foreach (var request in group)
                conflicts.Add(request.RowNumber);
        }
        return conflicts;
    }

    private static string CleanWhitespace(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : MultipleWhitespaceRegex().Replace(value.Normalize(NormalizationForm.FormKC).Trim(), " ");

    private static string NormalizeForComparison(string value) => CleanWhitespace(value).ToUpperInvariant();

    private static string CleanPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;
        var path = value.Normalize(NormalizationForm.FormKC).Trim().Trim('"').Replace('/', '\\');
        var isUnc = path.StartsWith(@"\\", StringComparison.Ordinal);
        var parts = path.Split('\\', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim()).Where(part => part.Length > 0);
        return (isUnc ? @"\\" : string.Empty) + string.Join('\\', parts);
    }

    private static bool IsValidFolderName(string name, out string reason)
    {
        if (name is "." or "..")
        {
            reason = "nome reservado";
            return false;
        }
        var invalidCharacter = name.FirstOrDefault(character => Path.GetInvalidFileNameChars().Contains(character));
        if (invalidCharacter != default)
        {
            reason = $"contém o caractere inválido '{invalidCharacter}'";
            return false;
        }
        if (name.EndsWith('.') || name.EndsWith(' '))
        {
            reason = "termina com ponto ou espaço";
            return false;
        }
        reason = string.Empty;
        return true;
    }

    private static bool IsSameOrInside(string candidate, string parent)
    {
        var fullCandidate = Path.TrimEndingDirectorySeparator(Path.GetFullPath(candidate));
        var fullParent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(parent));
        return fullCandidate.Equals(fullParent, StringComparison.OrdinalIgnoreCase) ||
               fullCandidate.StartsWith(fullParent + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }

    private static string? FindDefaultSpreadsheet()
    {
        var candidates = new List<string>
        {
            Path.Combine(Environment.CurrentDirectory, DefaultSpreadsheetName),
            Path.Combine(AppContext.BaseDirectory, DefaultSpreadsheetName)
        };
        AddParentCandidates(candidates, Environment.CurrentDirectory);
        AddParentCandidates(candidates, AppContext.BaseDirectory);
        return candidates.Distinct(StringComparer.OrdinalIgnoreCase).FirstOrDefault(File.Exists);
    }

    private static void AddParentCandidates(ICollection<string> candidates, string startPath)
    {
        var directory = new DirectoryInfo(startPath);
        for (var level = 0; level < 5 && directory is not null; level++, directory = directory.Parent)
            candidates.Add(Path.Combine(directory.FullName, DefaultSpreadsheetName));
    }

    private static void PrintSummary(ProcessingResult result, int spreadsheetRows)
    {
        Console.WriteLine();
        Console.WriteLine("Resumo:");
        Console.WriteLine($"  Linhas válidas na planilha:       {spreadsheetRows:N0}");
        Console.WriteLine($"  Pastas de origem correspondentes: {result.MatchedFolders:N0}");
        Console.WriteLine($"  Pastas não encontradas:            {result.SourceFoldersNotFound:N0}");
        Console.WriteLine($"  Linhas duplicadas ignoradas:       {result.DuplicateRows:N0}");
        Console.WriteLine($"  Arquivos processados:              {result.FilesProcessed:N0}");
        Console.WriteLine($"  Arquivos não sobrescritos:         {result.Collisions:N0}");
        Console.WriteLine($"  Erros:                             {result.Errors:N0}");
    }

    private static string WriteReport(
        string spreadsheetPath,
        string worksheetName,
        string sourceRoot,
        int spreadsheetRows,
        bool dryRun,
        ProcessingResult result)
    {
        Directory.CreateDirectory(ReportDirectory);
        var timestamp = DateTime.Now;
        var baseFileName = $"Relatorio_Copia_{timestamp:yyyyMMdd_HHmmss}";
        var reportPath = Path.Combine(ReportDirectory, baseFileName + ".txt");
        var suffix = 1;
        while (File.Exists(reportPath))
            reportPath = Path.Combine(ReportDirectory, $"{baseFileName}_{suffix++}.txt");

        var report = new StringBuilder();
        report.AppendLine("RELATÓRIO DE CÓPIA DE ARQUIVOS");
        report.AppendLine(new string('=', 80));
        report.AppendLine($"Data/hora: {timestamp:dd/MM/yyyy HH:mm:ss}");
        report.AppendLine($"Modo: {(dryRun ? "SIMULAÇÃO - nenhuma cópia foi realizada" : "EXECUÇÃO")}");
        report.AppendLine($"Planilha: {Path.GetFullPath(spreadsheetPath)}");
        report.AppendLine($"Aba: {worksheetName}");
        report.AppendLine($"Origem: {sourceRoot}");
        report.AppendLine($"Paralelismo: {result.Parallelism} reclamantes simultâneos");
        report.AppendLine($"Duração do processamento: {result.Elapsed.ToString(@"hh\:mm\:ss\.fff")}");
        report.AppendLine();
        report.AppendLine("RESUMO");
        report.AppendLine(new string('-', 80));
        report.AppendLine($"Linhas válidas na planilha: {spreadsheetRows:N0}");
        report.AppendLine($"Pastas correspondentes: {result.MatchedFolders:N0}");
        report.AppendLine($"Pastas não encontradas: {result.SourceFoldersNotFound:N0}");
        report.AppendLine($"Arquivos {(dryRun ? "que seriam copiados" : "copiados")}: {result.FilesProcessed:N0}");
        report.AppendLine($"Arquivos não sobrescritos: {result.Collisions:N0}");
        report.AppendLine($"Erros: {result.Errors:N0}");
        AppendReportSection(
            report,
            dryRun ? "CÓPIAS SIMULADAS" : "ARQUIVOS COPIADOS",
            result.CopiedFileDetails);
        AppendReportSection(report, "ERROS", result.ErrorDetails);
        AppendReportSection(report, "NÃO COPIADOS / AVISOS", result.WarningDetails);

        File.WriteAllText(reportPath, report.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        return reportPath;
    }

    private static void AppendReportSection(StringBuilder report, string title, IEnumerable<string> details)
    {
        var detailList = details.ToArray();
        report.AppendLine();
        report.AppendLine(title);
        report.AppendLine(new string('-', 80));
        if (detailList.Length == 0)
        {
            report.AppendLine("Nenhum registro.");
            return;
        }

        for (var index = 0; index < detailList.Length; index++)
            report.AppendLine($"{index + 1:N0}. {detailList[index]}");
    }

    private static void RegisterWarning(ProcessingResult result, string message)
    {
        Warning(message);
        result.WarningDetails.Enqueue(message);
    }

    private static void RegisterError(ProcessingResult result, string message)
    {
        Error(message);
        result.ErrorDetails.Enqueue(message);
        result.IncrementErrors();
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Uso: SeparacaoDeArquivosPorCompensacao [opções]");
        Console.WriteLine("  --planilha <arquivo.xlsx>  Planilha (padrão: procura PlanilhaPaths.xlsx).");
        Console.WriteLine("  --origem <pasta>           Pasta BOOKS com uma pasta por reclamante.");
        Console.WriteLine("  --simular                   Não cria pastas nem altera arquivos.");
        Console.WriteLine("  --paralelismo <número>      Reclamantes processados simultaneamente.");
        Console.WriteLine("  -h, --help                  Mostra esta ajuda.");
        Console.WriteLine();
        Console.WriteLine("Coluna A: reclamante. Coluna C: pasta-base de destino.");
        Console.WriteLine("O destino será <pasta-base>\\<reclamante>; arquivos existentes não serão sobrescritos.");
    }

    private static void Info(string message)
    {
        lock (ConsoleLock)
            Console.WriteLine(message);
    }

    private static void Warning(string message)
    {
        lock (ConsoleLock)
            Console.WriteLine($"[AVISO] {message}");
    }

    private static void Error(string message)
    {
        lock (ConsoleLock)
            Console.Error.WriteLine($"[ERRO] {message}");
    }

    [GeneratedRegex(@"\s+")]
    private static partial Regex MultipleWhitespaceRegex();

    private sealed record TransferRequest(int RowNumber, string Claimant, string ComparisonKey, string DestinationFolder);

    private sealed class ProcessingResult
    {
        private int _matchedFolders;
        private int _sourceFoldersNotFound;
        private int _duplicateRows;
        private int _filesProcessed;
        private int _collisions;
        private int _errors;

        public int MatchedFolders => Volatile.Read(ref _matchedFolders);
        public int SourceFoldersNotFound => Volatile.Read(ref _sourceFoldersNotFound);
        public int DuplicateRows => Volatile.Read(ref _duplicateRows);
        public int FilesProcessed => Volatile.Read(ref _filesProcessed);
        public int Collisions => Volatile.Read(ref _collisions);
        public int Errors => Volatile.Read(ref _errors);
        public int Parallelism { get; init; }
        public TimeSpan Elapsed { get; set; }
        public ConcurrentQueue<string> CopiedFileDetails { get; } = new();
        public ConcurrentQueue<string> ErrorDetails { get; } = new();
        public ConcurrentQueue<string> WarningDetails { get; } = new();

        public void IncrementMatchedFolders() => Interlocked.Increment(ref _matchedFolders);
        public void IncrementSourceFoldersNotFound() => Interlocked.Increment(ref _sourceFoldersNotFound);
        public void IncrementDuplicateRows() => Interlocked.Increment(ref _duplicateRows);
        public void IncrementCollisions() => Interlocked.Increment(ref _collisions);
        public void IncrementErrors() => Interlocked.Increment(ref _errors);

        public void AddCopiedFile(string detail)
        {
            CopiedFileDetails.Enqueue(detail);
            Interlocked.Increment(ref _filesProcessed);
        }
    }

    private sealed record AppOptions(
        string? SpreadsheetPath,
        string SourcePath,
        bool DryRun,
        bool ShowHelp,
        int Parallelism)
    {
        private static readonly int DefaultParallelism = Math.Clamp(Environment.ProcessorCount * 2, 4, 12);

        public static bool TryParse(string[] args, out AppOptions options, out string? error)
        {
            string? spreadsheetPath = null;
            var sourcePath = DefaultSourcePath;
            var dryRun = false;
            var showHelp = false;
            var parallelism = DefaultParallelism;

            for (var index = 0; index < args.Length; index++)
            {
                switch (args[index])
                {
                    case "-h" or "--help": showHelp = true; break;
                    case "--simular": dryRun = true; break;
                    case "--paralelismo":
                        if (!TryReadValue(args, ref index, out var parallelismText) ||
                            !int.TryParse(parallelismText, out parallelism) ||
                            parallelism <= 0 || parallelism > 64)
                        {
                            options = null!;
                            error = "A opção --paralelismo exige um número entre 1 e 64.";
                            return false;
                        }
                        break;
                    case "--planilha":
                        if (!TryReadValue(args, ref index, out spreadsheetPath))
                        {
                            options = null!; error = "A opção --planilha exige um caminho."; return false;
                        }
                        break;
                    case "--origem":
                        if (!TryReadValue(args, ref index, out sourcePath))
                        {
                            options = null!; error = "A opção --origem exige um caminho."; return false;
                        }
                        break;
                    default:
                        options = null!; error = $"Opção desconhecida: {args[index]}"; return false;
                }
            }

            options = new AppOptions(spreadsheetPath, sourcePath, dryRun, showHelp, parallelism);
            error = null;
            return true;
        }

        private static bool TryReadValue(string[] args, ref int index, out string value)
        {
            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                value = string.Empty;
                return false;
            }
            value = args[++index].Trim().Trim('"');
            return value.Length > 0;
        }
    }
}
