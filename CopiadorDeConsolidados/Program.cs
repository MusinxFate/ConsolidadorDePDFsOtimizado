using System.Collections.Concurrent;

namespace CopiadorDeConsolidados;

internal static class Program
{
    private const string DefaultSourcePath =
        @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\4. CÓPIAS DOS PROCESSOS - BOOKS";

    private const string DefaultDestinationPath =
        @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\13. PAFs\19613.7393632022-72 - 6ª COMPENSAÇÃO 07.2022\BOOK";



    private const int DefaultParallelism = 6;
    private static readonly object LogLock = new();

    private static async Task<int> Main(string[] args)
    {
        if (args.Length > 0 && args[0] is "-h" or "--help")
        {
            PrintUsage();
            return 0;
        }

        var sourcePath = Path.GetFullPath(args.ElementAtOrDefault(0) ?? DefaultSourcePath);
        var destinationPath = Path.GetFullPath(args.ElementAtOrDefault(1) ?? DefaultDestinationPath);

        if (!TryGetParallelism(args.ElementAtOrDefault(2), out var parallelism))
        {
            LogError("O paralelismo deve ser um número inteiro maior que zero.");
            PrintUsage();
            return 2;
        }

        if (!Directory.Exists(sourcePath))
        {
            LogError($"A pasta de origem não foi encontrada ou não está acessível: {sourcePath}");
            return 1;
        }

        if (PathsAreEqual(sourcePath, destinationPath) || IsInside(destinationPath, sourcePath))
        {
            LogError("A pasta de destino deve ficar fora da pasta de origem.");
            return 2;
        }

        using var cancellation = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellation.Cancel();
            LogWarning("Cancelamento solicitado. Finalizando as cópias em andamento...");
        };

        LogInfo($"Origem: {sourcePath}");
        LogInfo($"Destino: {destinationPath}");
        LogInfo($"Pastas selecionadas: {(SelectedFolderNames.Length == 0 ? "(nenhuma)" : string.Join(", ", SelectedFolderNames))}");
        LogInfo($"Cópias simultâneas: {parallelism}");

        try
        {
            Directory.CreateDirectory(destinationPath);
            LogInfo("Iniciando a localização e a cópia dos PDFs consolidados.");
            var result = await CopyConsolidatedPdfsAsync(
                sourcePath,
                destinationPath,
                parallelism,
                cancellation.Token);

            Console.WriteLine("\nConcluído:");
            Console.WriteLine($"  Encontrados: {result.Found:N0}");
            Console.WriteLine($"  Copiados:    {result.Copied:N0}");
            Console.WriteLine($"  Já iguais:   {result.Skipped:N0}");
            Console.WriteLine($"  Erros:       {result.Errors.Count:N0}");

            if (!result.Errors.IsEmpty)
            {
                Console.Error.WriteLine("\nArquivos que falharam:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  {error}");
            }

            return result.Errors.IsEmpty ? 0 : 1;
        }
        catch (OperationCanceledException)
        {
            LogWarning("Operação cancelada.");
            return 130;
        }
        catch (Exception ex)
        {
            LogError($"Falha durante o processamento das pastas: {ex.GetType().Name}: {ex.Message}");
            return 1;
        }
    }

    private static async Task<CopyResult> CopyConsolidatedPdfsAsync(
        string sourcePath,
        string destinationPath,
        int parallelism,
        CancellationToken cancellationToken)
    {
        var result = new CopyResult();
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = true,
            IgnoreInaccessible = true,
            MatchCasing = MatchCasing.CaseInsensitive,
            AttributesToSkip = FileAttributes.ReparsePoint,
            ReturnSpecialDirectories = false
        };
        var selectedFolderNames = new HashSet<string>(
            SelectedFolderNames,
            StringComparer.OrdinalIgnoreCase);

        // EnumerateFiles streams the results instead of loading all file names into memory.
        var files = Directory
            .EnumerateFiles(sourcePath, "*.pdf", options)
            .Where(path =>
                IsConsolidatedPdf(path) &&
                IsInSelectedFolder(path, sourcePath, selectedFolderNames));

        await Parallel.ForEachAsync(
            files,
            new ParallelOptions
            {
                MaxDegreeOfParallelism = parallelism,
                CancellationToken = cancellationToken
            },
            (sourceFile, token) =>
            {
                token.ThrowIfCancellationRequested();
                var found = Interlocked.Increment(ref result.Found);

                try
                {
                    var relativePath = Path.GetRelativePath(sourcePath, sourceFile);
                    var destinationFile = Path.Combine(destinationPath, relativePath);
                    var destinationDirectory = Path.GetDirectoryName(destinationFile)
                        ?? destinationPath;

                    Directory.CreateDirectory(destinationDirectory);

                    if (FilesAreEqual(sourceFile, destinationFile))
                    {
                        Interlocked.Increment(ref result.Skipped);
                        LogInfo($"Arquivo já está atualizado; cópia ignorada: '{sourceFile}' -> '{destinationFile}'");
                    }
                    else
                    {
                        LogInfo($"Copiando arquivo: '{sourceFile}' -> '{destinationFile}'");
                        File.Copy(sourceFile, destinationFile, overwrite: true);
                        Interlocked.Increment(ref result.Copied);
                        LogInfo($"Arquivo copiado com sucesso: '{destinationFile}'");
                    }

                    if (found % 100 == 0)
                        LogInfo($"Progresso: {found:N0} arquivos processados.");
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    var error = $"'{sourceFile}' -> {ex.GetType().Name}: {ex.Message}";
                    result.Errors.Enqueue(error);
                    LogError($"Falha ao processar o arquivo: {error}");
                }

                return ValueTask.CompletedTask;
            });

        return result;
    }

    private static bool IsConsolidatedPdf(string path) =>
        Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase) &&
        Path.GetFileNameWithoutExtension(path)
            .Contains("RT", StringComparison.OrdinalIgnoreCase);

    private static bool IsInSelectedFolder(
        string filePath,
        string sourcePath,
        HashSet<string> selectedFolderNames)
    {
        var relativeDirectory = Path.GetRelativePath(
            sourcePath,
            Path.GetDirectoryName(filePath)!);

        return relativeDirectory
            .Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Any(selectedFolderNames.Contains);
    }

    private static bool FilesAreEqual(string sourceFile, string destinationFile)
    {
        var destinationInfo = new FileInfo(destinationFile);
        if (!destinationInfo.Exists)
            return false;

        var sourceInfo = new FileInfo(sourceFile);
        return sourceInfo.Length == destinationInfo.Length &&
               sourceInfo.LastWriteTimeUtc == destinationInfo.LastWriteTimeUtc;
    }

    private static bool TryGetParallelism(string? value, out int parallelism)
    {
        if (value is null)
        {
            parallelism = DefaultParallelism;
            return true;
        }

        return int.TryParse(value, out parallelism) && parallelism > 0;
    }

    private static bool PathsAreEqual(string firstPath, string secondPath) =>
        string.Equals(
            Path.TrimEndingDirectorySeparator(firstPath),
            Path.TrimEndingDirectorySeparator(secondPath),
            StringComparison.OrdinalIgnoreCase);

    private static bool IsInside(string candidatePath, string parentPath)
    {
        var parentWithSeparator = Path.TrimEndingDirectorySeparator(parentPath) + Path.DirectorySeparatorChar;
        return candidatePath.StartsWith(parentWithSeparator, StringComparison.OrdinalIgnoreCase);
    }

    private static void LogInfo(string message) => Log("INFO", message, isError: false);

    private static void LogWarning(string message) => Log("AVISO", message, isError: false);

    private static void LogError(string message) => Log("ERRO", message, isError: true);

    private static void Log(string level, string message, bool isError)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";

        // Evita que mensagens de cópias paralelas sejam misturadas no console.
        lock (LogLock)
        {
            if (isError)
                Console.Error.WriteLine(line);
            else
                Console.WriteLine(line);
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Uso:");
        Console.WriteLine("  CopiadorDeConsolidados [origem] [destino] [copias-simultaneas]");
        Console.WriteLine();
        Console.WriteLine("Sem argumentos, o programa usa os caminhos de rede configurados no código.");
    }

    private sealed class CopyResult
    {
        public int Found;
        public int Copied;
        public int Skipped;
        public ConcurrentQueue<string> Errors { get; } = new();
    }
    
    private static readonly string[] SelectedFolderNames = [

    ];
}
