using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;

[assembly: SupportedOSPlatform("windows")]

namespace ConsolidacaoDePDFs
{
    internal class Program
    {
        private static string LogFolderPath = @"C:\Users\murilo.souza\Desktop\Testes - Decadencia\Logs\";
        private const string PendingPublicationName = ".consolidacao-pendente";
        private const string PublicationTempPrefix = ".consolidacao-";
        private static readonly object _lock = new object();
        private const long MaxConsolidatedPdfSizeBytes = (long)(14.5 * 1024 * 1024);

        private const string GhostscriptMissingLogMessage =
            "Ghostscript não encontrado. Não é possível reduzir uma página única acima de 14,5 MB apenas com split.";

        private static readonly Lazy<string?> GhostscriptExecutablePath = new(FindGhostscriptExecutable);

        private static readonly string BasePath =
        //@"\\SRV-FS02-BBM\Empresas$\Empresas\AMCOR\1. CONTROLE\3. CP NAS RTS- DECADÊNCIA\3. BOOKS";
        //@"C:\Users\murilo.souza\Desktop\Testes - Decadencia\Fazer";
        @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\13. PAFs\19613.7393632022-72 - 6ª COMPENSAÇÃO 07.2022\BOOK";

        static int Main(string[] args)
        {
            var basePath = BasePath;
            string? destination = null;
            int? requestedParallelism = null;
            var copyInputs = false;

            foreach (var argument in args)
            {
                if (argument == "--copiar-entradas")
                    copyInputs = true;
                else if (argument.StartsWith("--paralelismo=") &&
                         int.TryParse(argument["--paralelismo=".Length..], out var value) && value > 0)
                    requestedParallelism = value;
                else if (argument.StartsWith("--pasta=") && argument.Length > "--pasta=".Length)
                    basePath = argument["--pasta=".Length..];
                else if (argument.StartsWith("--logs=") && argument.Length > "--logs=".Length)
                    LogFolderPath = argument["--logs=".Length..];
                else if (argument.StartsWith("--destino=") &&
                         !string.IsNullOrWhiteSpace(argument["--destino=".Length..]))
                    destination = Path.TrimEndingDirectorySeparator(Path.GetFullPath(argument["--destino=".Length..]));
                else
                {
                    Console.Error.WriteLine($"Argumento inválido: {argument}");
                    Console.Error.WriteLine("Uso: [--paralelismo=N] [--copiar-entradas] [--pasta=CAMINHO] [--logs=CAMINHO] [--destino=CAMINHO]");
                    return 1;
                }
            }

            if (Console.IsInputRedirected)
                Console.InputEncoding = Encoding.UTF8;

            if (ReadYesNo("Deseja dividir os PDFs consolidados com base no tamanho? (S/N): ")
                is not bool dividirPdfsPorTamanho)
            {
                Console.Error.WriteLine("Entrada encerrada antes da escolha sobre a divisão dos PDFs.");
                return 1;
            }

            destination ??= ReadDestination();
            var parallelism = requestedParallelism ?? ReadParallelism(Math.Min(6, Environment.ProcessorCount));
            copyInputs = copyInputs || ReadYesNo("Deseja copiar os PDFs de entrada para o disco local? (S/N) [N]: ", false) == true;

            var sw = Stopwatch.StartNew();
            Directory.CreateDirectory(LogFolderPath);

            Log(dividirPdfsPorTamanho
                ? "Divisão dos PDFs por tamanho: SIM"
                : "Divisão dos PDFs por tamanho: NÃO (somente consolidação)");

            var nomes = GetNomes();

            var nomesSet = new HashSet<string>(nomes);

            Log($"Pastas simultâneas: {parallelism}; cópia local dos PDFs de entrada: {(copyInputs ? "SIM" : "NÃO")}");
            var folders = Directory.EnumerateDirectories(Path.GetFullPath(basePath));

            var matchingFolders = new List<string>();
            var encontrados = new HashSet<string>();

            var nomesPastas = new List<string>();

            foreach (var folder in folders)
            {
                var folderName = Path.GetFileName(folder);
                if (string.Equals(folder, destination, StringComparison.OrdinalIgnoreCase))
                {
                    if (nomesSet.Contains(folderName))
                    {
                        Console.Error.WriteLine($"O destino deve ser diferente da pasta de entrada: {folder}");
                        return 1;
                    }

                    continue;
                }

                nomesPastas.Add(folderName);

                if (nomesSet.Contains(folderName))
                {
                    matchingFolders.Add(folder);
                    encontrados.Add(folderName);
                }
            }

            if (destination is not null)
            {
                Directory.CreateDirectory(destination);
                Log($"Destino dos PDFs consolidados: {destination}");
            }

            var naoEncontrados = nomesSet
                .Where(n => !encontrados.Contains(n))
                .ToList();

            File.WriteAllLines(Path.Combine(LogFolderPath, "consolidacao_log_existe_excel_nao_existe_pasta.txt"), naoEncontrados);

            var pastasNaoNaLista = nomesPastas
                .Where(nome => !nomesSet.Contains(nome))
                .ToList();

            File.WriteAllLines(Path.Combine(LogFolderPath, "consolidacao_log_existe_pasta_nao_existe_no_excel.txt"), pastasNaoNaLista);

            Log($"Pastas para processar: {matchingFolders.Count}");
            Log($"Não encontrados: {naoEncontrados.Count}");
            Log($"Pastas fora da lista: {pastasNaoNaLista.Count}");

            var options = new ParallelOptions
            {
                MaxDegreeOfParallelism = parallelism
            };

            using (var word = new WordConverter())
            {
                Parallel.ForEach(matchingFolders, options,
                    folder => ProcessFolder(folder, destination, dividirPdfsPorTamanho, copyInputs, word));
            }

            sw.Stop();
            Log($"FINALIZADO em {sw.Elapsed}");
            return 0;
        }

        static bool? ReadYesNo(string prompt, bool? defaultAnswer = null)
        {
            while (true)
            {
                Console.Write(prompt);
                var answer = Console.ReadLine()?.Trim();
                if (answer is null || (answer.Length == 0 && defaultAnswer.HasValue))
                    return defaultAnswer;

                if (answer.Equals("S", StringComparison.OrdinalIgnoreCase))
                    return true;

                if (answer.Equals("N", StringComparison.OrdinalIgnoreCase))
                    return false;

                Console.WriteLine("Opção inválida. Digite S para sim ou N para não.");
            }
        }

        static string? ReadDestination()
        {
            while (true)
            {
                Console.Write("Pasta única para salvar os consolidados [Enter para usar as pastas de origem]: ");
                var path = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(path))
                    return null;

                try
                {
                    if (path.Length >= 2 && path[0] == '"' && path[^1] == '"')
                        path = path[1..^1];

                    return Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
                }
                catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
                {
                    Console.WriteLine($"Caminho inválido. {ex.Message}");
                }
            }
        }

        static int ReadParallelism(int defaultParallelism)
        {
            while (true)
            {
                Console.Write($"Quantas pastas deseja processar simultaneamente? [{defaultParallelism}]: ");
                var answer = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(answer))
                    return defaultParallelism;

                if (int.TryParse(answer, out var parallelism) && parallelism > 0)
                    return parallelism;

                Console.WriteLine("Digite um número inteiro maior que zero.");
            }
        }

        static void ProcessFolder(string folder, string? destination, bool dividirPdfsPorTamanho, bool copyInputs,
            WordConverter word)
        {
            DirectoryInfo? workDirectory = null;

            try
            {
                var timer = Stopwatch.StartNew();
                var folderName = Path.GetFileName(folder);
                var outputFolder = destination ?? folder;
                var outputName = $"RT {folderName}.pdf";
                var outputPath = Path.Combine(outputFolder, outputName);
                var tempPrefix = destination is null ? PublicationTempPrefix : GetPublicationTempPrefix(outputName);
                var pendingName = destination is null ? PendingPublicationName : $"{tempPrefix}pendente";
                var lockName = destination is null ? ".consolidacao.lock" : $"{tempPrefix}lock";
                var outputFiles = ReadOutputFiles(outputFolder, folderName, pendingName, tempPrefix, destination is not null);
                var listTime = timer.Elapsed;

                var existingPdf = FindFile(outputFiles, outputName);
                var hasPendingPublication = FindFile(outputFiles, pendingName) is not null;
                var staleUploads = outputFiles.Where(file => IsPublicationTemporaryFile(file.Name, tempPrefix)).ToArray();
                var hasInterruptedWork = hasPendingPublication || staleUploads.Length > 0;
                if (!hasInterruptedWork && existingPdf is not null &&
                    (!dividirPdfsPorTamanho || existingPdf.Length <= MaxConsolidatedPdfSizeBytes))
                {
                    Log($"[SKIP] {folderName}");
                    LogJaExisteConsolidacao_(folderName);
                    return;
                }

                var parts = existingPdf is null ? GetConsolidatedParts(outputFiles, folderName) : [];
                var pendingPath = Path.Combine(outputFolder, pendingName);
                if (!hasInterruptedWork && parts.Length > 0 &&
                    (!dividirPdfsPorTamanho || parts.All(part => part.Length <= MaxConsolidatedPdfSizeBytes)) &&
                    !File.Exists(pendingPath))
                {
                    Log($"[SKIP PARTS] {folderName}");
                    LogJaExisteConsolidacao_(folderName);
                    return;
                }

                FileInfo[]? files = destination is null ? outputFiles : null;
                if (!hasInterruptedWork && existingPdf is null && parts.Length == 0)
                {
                    files ??= ReadFolder(folder);
                    listTime = timer.Elapsed;
                    if (!files.Any(file => file.Length > 0 && IsInputFile(file, folderName)))
                    {
                        Log($"[EMPTY] {folderName}");
                        LogPastasVazias(folderName);
                        return;
                    }
                }

                using var folderLock = new FileStream(Path.Combine(outputFolder, lockName),
                    FileMode.OpenOrCreate, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose);

                foreach (var upload in staleUploads)
                    DeleteTemporaryFile(upload.FullName);

                var pendingPublicationStillExists = File.Exists(pendingPath);
                if (pendingPublicationStillExists)
                    RecoverPendingPublication(outputFolder, folderName, pendingPath);

                if (hasInterruptedWork || pendingPublicationStillExists)
                {
                    outputFiles = ReadOutputFiles(outputFolder, folderName, pendingName, tempPrefix, destination is not null);
                    if (destination is null)
                        files = outputFiles;

                    existingPdf = FindFile(outputFiles, outputName);
                    parts = existingPdf is null ? GetConsolidatedParts(outputFiles, folderName) : [];
                    listTime = timer.Elapsed;
                }

                workDirectory = Directory.CreateTempSubdirectory("consolidacao-pdf-");
                var workPath = workDirectory.FullName;
                var splitTime = TimeSpan.Zero;
                var publishTime = TimeSpan.Zero;

                if (existingPdf is not null)
                {
                    if (dividirPdfsPorTamanho && existingPdf.Length > MaxConsolidatedPdfSizeBytes)
                        SplitExistingPdf(existingPdf, workPath, tempPrefix, ref splitTime, ref publishTime);

                    Log($"[SKIP] {folderName} | listagem {listTime.TotalSeconds:F2}s | divisão {splitTime.TotalSeconds:F2}s | publicação {publishTime.TotalSeconds:F2}s");
                    LogJaExisteConsolidacao_(folderName);
                    return;
                }

                if (parts.Length > 0)
                {
                    if (dividirPdfsPorTamanho)
                    {
                        foreach (var part in parts.Where(file => file.Length > MaxConsolidatedPdfSizeBytes))
                            SplitExistingPdf(part, workPath, tempPrefix, ref splitTime, ref publishTime);
                    }

                    Log($"[SKIP PARTS] {folderName} | listagem {listTime.TotalSeconds:F2}s | divisão {splitTime.TotalSeconds:F2}s | publicação {publishTime.TotalSeconds:F2}s");
                    LogJaExisteConsolidacao_(folderName);
                    return;
                }

                if (File.Exists(outputPath) || File.Exists(Path.Combine(outputFolder, GetPartName(outputName, 1))))
                {
                    Log($"[SKIP CONCORRENTE] {folderName}");
                    LogJaExisteConsolidacao_(folderName);
                    return;
                }

                if (files is null)
                {
                    files = ReadFolder(folder);
                    listTime = timer.Elapsed;
                }

                var phase = Stopwatch.StartNew();
                var converted = ConvertDocFilesToPdf(folder, folderName, files, workPath, word,
                    publishToSource: destination is null);
                var wordTime = phase.Elapsed;
                var inputs = files.Concat(converted)
                    .Where(file => IsPdf(file) && file.Length > 0 && !IsConsolidatedPdf(file.Name, folderName))
                    .OrderBy(file => file.Name)
                    .ToArray();

                if (inputs.Length == 0)
                {
                    Log($"[EMPTY] {folderName}");
                    LogPastasVazias(folderName);
                    return;
                }

                phase.Restart();
                var localOutput = Path.Combine(workPath, outputName);
                var localInput = copyInputs ? Path.Combine(workPath, "entrada.tmp") : null;
                using (var output = CreatePdfDocument())
                {
                    foreach (var input in inputs)
                    {
                        var isSourceInput = string.Equals(input.DirectoryName, folder, StringComparison.OrdinalIgnoreCase);
                        TryAppendPdf(input.FullName, isSourceInput ? localInput : null, output);
                    }

                    if (output.PageCount == 0)
                    {
                        Log($"[NO PAGES] {folderName}");
                        return;
                    }

                    output.Save(localOutput);
                }

                var mergeTime = phase.Elapsed;
                phase.Restart();
                var filesToPublish = new List<string> { localOutput };
                if (dividirPdfsPorTamanho && GetFileSize(localOutput) > MaxConsolidatedPdfSizeBytes)
                {
                    var split = SplitPdf(localOutput, outputPath);
                    if (split.FitsLimit)
                        filesToPublish.Clear();

                    filesToPublish.AddRange(split.Parts);
                }

                splitTime = phase.Elapsed;
                phase.Restart();
                var journal = filesToPublish.Count > 1 && filesToPublish[0] != localOutput ? pendingPath : null;
                PublishFiles(filesToPublish, outputFolder, overwrite: false, journal, tempPrefix);
                publishTime = phase.Elapsed;
                Log($"[OK] {folderName} | {inputs.Length} PDFs | listagem {listTime.TotalSeconds:F2}s | Word {wordTime.TotalSeconds:F2}s | consolidação {mergeTime.TotalSeconds:F2}s | divisão {splitTime.TotalSeconds:F2}s | publicação {publishTime.TotalSeconds:F2}s");
            }
            catch (Exception ex)
            {
                Log($"[ERRO PASTA] {folder} => {ex.Message}");
            }
            finally
            {
                if (workDirectory is not null)
                {
                    try
                    {
                        workDirectory.Delete(recursive: true);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        Log($"[ERRO TEMPORÁRIO] {workDirectory.FullName} => {ex.Message}");
                    }
                }
            }
        }

        static FileInfo[] ReadFolder(string folder)
        {
            return new DirectoryInfo(folder).EnumerateFiles()
                .Where(file => IsPdf(file) || IsWordDocument(file.Name) || file.Name == PendingPublicationName ||
                               IsPublicationTemporaryFile(file.Name))
                .ToArray();
        }

        static FileInfo[] ReadOutputFiles(string folder, string folderName, string pendingName, string tempPrefix,
            bool sharedDestination)
        {
            if (!sharedDestination)
                return ReadFolder(folder);

            var directory = new DirectoryInfo(folder);
            return directory.EnumerateFiles($"RT {folderName}*").Where(IsPdf)
                .Concat(directory.EnumerateFiles($"{tempPrefix}*")
                    .Where(file => file.Name == pendingName || IsPublicationTemporaryFile(file.Name, tempPrefix)))
                .ToArray();
        }

        static string GetPublicationTempPrefix(string outputName)
        {
            var hash = SHA256.HashData(Encoding.UTF8.GetBytes(outputName.ToUpperInvariant()));
            return $"{PublicationTempPrefix}{Convert.ToHexString(hash)}-";
        }

        static bool IsPublicationTemporaryFile(string name, string tempPrefix = PublicationTempPrefix)
        {
            return name.Length == tempPrefix.Length + 36 &&
                   name.StartsWith(tempPrefix, StringComparison.Ordinal) &&
                   name.EndsWith(".tmp", StringComparison.Ordinal) &&
                   Guid.TryParseExact(name.AsSpan(tempPrefix.Length, 32), "N", out _);
        }

        static FileInfo? FindFile(FileInfo[] files, string name)
        {
            return files.FirstOrDefault(file => file.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        static bool IsPdf(FileInfo file)
        {
            return file.Extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsInputFile(FileInfo file, string folderName)
        {
            return IsWordDocument(file.Name) ||
                   (IsPdf(file) && !IsConsolidatedPdf(file.Name, folderName));
        }

        static IReadOnlyList<FileInfo> ConvertDocFilesToPdf(string folder, string folderName, FileInfo[] files,
            string workPath, WordConverter word, bool publishToSource)
        {
            var documents = files
                .Where(file => IsWordDocument(file.Name) && !IsConsolidatedPdf(file.Name, folderName) && file.Length > 0)
                .OrderBy(file => file.Name)
                .ToArray();

            if (documents.Length == 0)
                return [];

            var pdfNames = files.Where(IsPdf).Select(file => file.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var toConvert = documents.Where(file => !pdfNames.Contains(Path.ChangeExtension(file.Name, ".pdf"))).ToArray();
            if (toConvert.Length == 0)
                return [];

            var localWordPath = Directory.CreateDirectory(Path.Combine(workPath, "word")).FullName;
            var exported = new List<(string Source, string Output)>();
            try
            {
                word.Run(application =>
                {
                    foreach (var file in toConvert)
                    {
                        var pdfName = Path.ChangeExtension(file.Name, ".pdf");
                        if (pdfNames.Contains(pdfName))
                            continue;

                        var localPdf = Path.Combine(localWordPath, pdfName);

                        try
                        {
                            ExportDocumentToPdf(application, file.FullName, localPdf);
                            pdfNames.Add(pdfName);
                            exported.Add((file.FullName, localPdf));
                        }
                        catch (Exception ex)
                        {
                            Log($"[ERRO DOC] {file.FullName} => {ex.Message}");
                            DeleteTemporaryFile(localPdf);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Log($"[ERRO WORD] {folder} => {ex.Message}");
            }

            var converted = new List<FileInfo>(exported.Count);
            foreach (var (source, localPdf) in exported)
            {
                try
                {
                    if (publishToSource)
                        PublishFiles(new[] { localPdf }, folder, overwrite: false);
                    converted.Add(new FileInfo(localPdf));
                    Log($"[DOC->PDF] {source}");
                }
                catch (Exception ex)
                {
                    Log($"[ERRO DOC] {source} => {ex.Message}");
                }
            }

            return converted;
        }

        static bool IsWordDocument(string file)
        {
            var extension = Path.GetExtension(file);
            return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
                   || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase);
        }

        static void ExportDocumentToPdf(dynamic word, string docPath, string pdfPath)
        {
            object? documents = null;
            object? document = null;

            try
            {
                documents = word.Documents;
                document = ((dynamic)documents).Open(
                    FileName: docPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
                ((dynamic)document).SaveAs2(FileName: pdfPath, FileFormat: 17);
            }
            finally
            {
                try
                {
                    if (document is not null)
                        ((dynamic)document).Close(SaveChanges: false);
                }
                finally
                {
                    ReleaseComObject(document);
                    ReleaseComObject(documents);
                }
            }
        }

        private sealed class WordConverter : IDisposable
        {
            private readonly BlockingCollection<(Action<object> Convert, TaskCompletionSource Completed)> _jobs = new();
            private readonly Lazy<Thread> _thread;

            public WordConverter()
            {
                _thread = new Lazy<Thread>(() =>
                {
                    var thread = new Thread(ProcessJobs);
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    return thread;
                });
            }

            public void Run(Action<object> convert)
            {
                var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                _ = _thread.Value;
                _jobs.Add((convert, completed));
                completed.Task.GetAwaiter().GetResult();
            }

            private void ProcessJobs()
            {
                object? application = null;

                try
                {
                    foreach (var (convert, completed) in _jobs.GetConsumingEnumerable())
                    {
                        try
                        {
                            if (application is null)
                            {
                                var wordType = Type.GetTypeFromProgID("Word.Application")
                                               ?? throw new InvalidOperationException("Microsoft Word indisponível para automação COM.");
                                application = Activator.CreateInstance(wordType)
                                              ?? throw new InvalidOperationException("Não foi possível iniciar o Microsoft Word.");
                                ((dynamic)application).Visible = false;
                                ((dynamic)application).DisplayAlerts = 0;
                            }

                            convert(application);
                            completed.SetResult();
                        }
                        catch (Exception ex)
                        {
                            completed.SetException(ex);
                        }
                    }
                }
                finally
                {
                    if (application is not null)
                    {
                        try
                        {
                            ((dynamic)application).Quit(SaveChanges: false);
                        }
                        catch (COMException ex)
                        {
                            Log($"[ERRO WORD] Encerramento => {ex.Message}");
                        }
                        finally
                        {
                            ReleaseComObject(application);
                        }
                    }
                }
            }

            public void Dispose()
            {
                _jobs.CompleteAdding();
                if (_thread.IsValueCreated)
                    _thread.Value.Join();

                _jobs.Dispose();
            }
        }

        static void SplitExistingPdf(FileInfo source, string workPath, string tempPrefix, ref TimeSpan splitTime,
            ref TimeSpan publishTime)
        {
            var timer = Stopwatch.StartNew();
            var splitDirectory = Directory.CreateDirectory(Path.Combine(workPath, Guid.NewGuid().ToString("N")));
            var localSource = Path.Combine(splitDirectory.FullName, source.Name);

            if (!File.Exists(source.FullName))
                return;

            CopyToLocal(source.FullName, localSource);

            if (GetFileSize(localSource) <= MaxConsolidatedPdfSizeBytes)
                return;

            var split = SplitPdf(localSource, source.FullName);
            splitTime += timer.Elapsed;
            timer.Restart();

            PublishFiles(split.Parts, source.DirectoryName!, overwrite: true, tempPrefix: tempPrefix);
            if (split.FitsLimit)
                File.Delete(source.FullName);

            publishTime += timer.Elapsed;
        }

        static (List<string> Parts, bool FitsLimit) SplitPdf(string localPdfPath, string destinationPath)
        {
            using var input = OpenWithRetry(localPdfPath);

            if (input.PageCount == 0)
                throw new InvalidDataException($"O PDF não tem páginas: {destinationPath}");

            var localDirectory = Path.GetDirectoryName(localPdfPath)!;
            var candidatePath = Path.Combine(localDirectory, "intervalo.tmp");
            var originalName = Path.GetFileName(localPdfPath);
            var parts = new List<string>();
            var fitsLimit = true;
            var startPageIndex = 0;

            while (startPageIndex < input.PageCount)
            {
                var partName = GetPartName(originalName, parts.Count + 1);
                var partPath = Path.Combine(localDirectory, partName);
                var rangeEnd = FindLargestPageRangeEnd(input, startPageIndex, candidatePath, partPath);

                if (rangeEnd < startPageIndex)
                {
                    File.Move(candidatePath, partPath, overwrite: true);
                    if (!TryCompressPdfToLimit(partPath))
                    {
                        fitsLimit = false;
                        Log($"[SPLIT ERRO LIMITE] {destinationPath} | página {startPageIndex + 1} | {partName} ({FormatSize(GetFileSize(partPath))})");
                    }

                    rangeEnd = startPageIndex;
                }

                parts.Add(partPath);
                startPageIndex = rangeEnd + 1;
            }

            Log($"[SPLIT] {destinationPath} => {parts.Count} partes");
            return (parts, fitsLimit);
        }

        static void PublishFiles(IReadOnlyList<string> localFiles, string folder, bool overwrite,
            string? journalPath = null, string tempPrefix = PublicationTempPrefix)
        {
            var staged = new List<(string TempPath, string Destination)>(localFiles.Count);
            var published = new List<string>(localFiles.Count);
            string? journalTempPath = null;
            var journalCreated = false;

            try
            {
                foreach (var localFile in localFiles)
                {
                    var destination = Path.Combine(folder, Path.GetFileName(localFile));
                    var tempPath = Path.Combine(folder, $"{tempPrefix}{Guid.NewGuid():N}.tmp");
                    staged.Add((tempPath, destination));
                    File.Copy(localFile, tempPath);
                }

                if (journalPath is not null)
                {
                    foreach (var (_, destination) in staged)
                    {
                        if (File.Exists(destination))
                            throw new IOException($"O arquivo já existe: {destination}");
                    }

                    journalTempPath = Path.Combine(folder, $"{tempPrefix}{Guid.NewGuid():N}.tmp");
                    File.WriteAllLines(journalTempPath, staged.Select(file => Path.GetFileName(file.Destination)));
                    File.Move(journalTempPath, journalPath);
                    journalTempPath = null;
                    journalCreated = true;
                }

                foreach (var (tempPath, destination) in staged)
                {
                    File.Move(tempPath, destination, overwrite);
                    published.Add(destination);
                }
            }
            catch
            {
                if (!overwrite)
                {
                    foreach (var destination in published)
                        File.Delete(destination);

                    if (journalCreated)
                        File.Delete(journalPath!);
                }

                throw;
            }
            finally
            {
                foreach (var (tempPath, _) in staged.Skip(published.Count))
                    DeleteTemporaryFile(tempPath);

                if (journalTempPath is not null)
                    DeleteTemporaryFile(journalTempPath);
            }

            if (journalCreated)
                File.Delete(journalPath!);
        }

        static void RecoverPendingPublication(string folder, string folderName, string journalPath)
        {
            var names = File.ReadAllLines(journalPath);
            var outputName = $"RT {folderName}.pdf";
            if (names.Length < 2 || names.Where((name, index) =>
                    !name.Equals(GetPartName(outputName, index + 1), StringComparison.Ordinal)).Any())
            {
                throw new InvalidDataException($"Registro de publicação inválido: {journalPath}");
            }

            var destinations = names.Select(name => Path.Combine(folder, name)).ToArray();
            if (!destinations.All(File.Exists))
            {
                foreach (var destination in destinations)
                    File.Delete(destination);

                Log($"[RECUPERADO] Publicação incompleta removida em {folder}");
            }

            File.Delete(journalPath);
        }

        static void DeleteTemporaryFile(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Log($"[ERRO TEMPORÁRIO] {path} => {ex.Message}");
            }
        }

        static bool TryCompressPdfToLimit(string pdfPath)
        {
            if (GetFileSize(pdfPath) <= MaxConsolidatedPdfSizeBytes)
            {
                return true;
            }

            var ghostscriptPath = GhostscriptExecutablePath.Value;

            if (ghostscriptPath is null)
            {
                Log($"[COMPRESS SKIP] {GhostscriptMissingLogMessage} Arquivo: {pdfPath}");
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

                var errorOutput = process.StandardError.ReadToEndAsync();
                var standardOutput = process.StandardOutput.BaseStream.CopyToAsync(Stream.Null);
                process.WaitForExit();
                var error = errorOutput.GetAwaiter().GetResult();
                standardOutput.GetAwaiter().GetResult();

                if (process.ExitCode != 0 || !File.Exists(tempPath))
                {
                    Log($"[COMPRESS ERRO] {pdfPath} ({profile}) => {error}");
                    continue;
                }

                var originalSize = GetFileSize(pdfPath);
                var compressedSize = GetFileSize(tempPath);

                if (compressedSize > 0 && compressedSize < originalSize)
                {
                    File.Move(tempPath, pdfPath, overwrite: true);
                    Log($"[COMPRESS] {pdfPath} {FormatSize(originalSize)} => {FormatSize(compressedSize)} ({profile})");
                }
                else
                {
                    File.Delete(tempPath);
                }

                if (GetFileSize(pdfPath) <= MaxConsolidatedPdfSizeBytes)
                {
                    return true;
                }
            }

            return false;
        }

        static string? FindGhostscriptExecutable()
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

                foreach (var candidate in Directory.EnumerateFiles(ghostscriptFolder, "gswin*c.exe",
                             SearchOption.AllDirectories))
                {
                    return candidate;
                }
            }

            return null;
        }

        static int FindLargestPageRangeEnd(PdfDocument input, int startPageIndex, string candidatePath,
            string partPath)
        {
            var lastPageIndex = input.PageCount - 1;
            var best = startPageIndex - 1;
            var candidateEnd = startPageIndex;

            // Aumenta o intervalo aos poucos para evitar gerar candidatos muito maiores que o limite.
            while (SaveCandidateIfItFits(input, startPageIndex, candidateEnd, candidatePath, partPath))
            {
                best = candidateEnd;
                if (best == lastPageIndex)
                    return best;

                var pages = best - startPageIndex + 1L;
                candidateEnd = (int)Math.Min(lastPageIndex, startPageIndex + 2 * pages - 1);
            }

            var low = best + 1;
            var high = candidateEnd - 1;

            while (low <= high)
            {
                var mid = low + (high - low) / 2;
                if (SaveCandidateIfItFits(input, startPageIndex, mid, candidatePath, partPath))
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

        static bool SaveCandidateIfItFits(PdfDocument input, int first, int last, string candidatePath,
            string partPath)
        {
            using (var output = CreatePdfDocument())
            {
                for (var index = first; index <= last; index++)
                    output.AddPage(input.Pages[index]);

                output.Save(candidatePath);
            }

            if (GetFileSize(candidatePath) > MaxConsolidatedPdfSizeBytes)
                return false;

            File.Move(candidatePath, partPath, overwrite: true);
            return true;
        }

        static PdfDocument CreatePdfDocument()
        {
            var document = new PdfDocument();
            ConfigurePdfOptions(document);
            return document;
        }

        static void ConfigurePdfOptions(PdfDocument document)
        {
            document.Options.CompressContentStreams = true;
            document.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            document.Options.UseFlateDecoderForJpegImages = PdfUseFlateDecoderForJpegImages.Automatic;
            document.Options.EnableCcittCompressionForBilevelImages = true;
        }

        static long GetFileSize(string filePath)
        {
            return new FileInfo(filePath).Length;
        }

        static string FormatSize(long bytes)
        {
            return $"{bytes / 1024d / 1024d:0.00} MB";
        }

        static string GetPartName(string pdfName, int partNumber)
        {
            return $"{Path.GetFileNameWithoutExtension(pdfName)} - PARTE {partNumber:00}.pdf";
        }

        static FileInfo[] GetConsolidatedParts(IEnumerable<FileInfo> files, string folderName)
        {
            var currentPrefix = $"RT {folderName} - PARTE ";
            var legacyPrefix = $"RT {folderName} PARTE ";
            return files
                .Where(file => IsPdf(file) &&
                               (file.Name.StartsWith(currentPrefix, StringComparison.OrdinalIgnoreCase) ||
                                file.Name.StartsWith(legacyPrefix, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(file => file.Name)
                .ToArray();
        }

        static bool IsConsolidatedPdf(string file, string folderName)
        {
            var fileName = Path.GetFileNameWithoutExtension(file);
            return fileName.StartsWith($"RT {folderName}", StringComparison.OrdinalIgnoreCase);
        }

        static void ReleaseComObject(object? comObject)
        {
            if (comObject is not null && Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }
        }

        static void TryAppendPdf(string file, string? localInput, PdfDocument output)
        {
            try
            {
                using var input = OpenWithRetry(file, localInput);
                CopyPages(input, output);
            }
            catch (Exception ex)
            {
                Log($"[ERRO ARQUIVO] {file} => {ex.Message}");
            }
        }

        static void CopyToLocal(string sourcePath, string localPath)
        {
            using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.Read, 1, FileOptions.SequentialScan);
            using var target = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None, 1);
            source.CopyTo(target, 128 * 1024);
        }

        static PdfDocument OpenWithRetry(string file, string? localInput = null)
        {
            const int attempts = 3;
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    if (localInput is not null)
                        CopyToLocal(file, localInput);

                    return PdfReader.Open(localInput ?? file, PdfDocumentOpenMode.Import);
                }
                catch (IOException ex) when (attempt < attempts && IsTransientIoError(ex))
                {
                    Thread.Sleep(200 * attempt);
                }
            }
        }

        static bool IsTransientIoError(IOException exception)
        {
            return (exception.HResult & 0xffff) is 32 or 33 or 53 or 54 or 58 or 59 or 64 or 67 or 121 or 1231 or 1232 or 1236;
        }

        static void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }

        static void Log(string msg)
        {
            var linha = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                Console.WriteLine(linha);
                File.AppendAllText(Path.Combine(LogFolderPath, "consolidacao_log_infos.txt"), linha + Environment.NewLine);
            }
        }

        static void LogPastasVazias(string msg)
        {
            var linha = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                File.AppendAllText(Path.Combine(LogFolderPath, "consolidacao_log_pastas_vazias.txt"), linha + Environment.NewLine);
            }
        }

        static void LogJaExisteConsolidacao_(string msg)
        {
            var linha = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                File.AppendAllText(Path.Combine(LogFolderPath, "consolidacao_log_ja_consolidado.txt"), linha + Environment.NewLine);
            }
        }

        static string[] GetNomes()
        {
            return new string[]
            {
"DANILO NUNES RODRIGUES",
"JOSE CARLOS MESQUITA PORTELA",
"DIEGO COSTA DE SANTANA",
"FLAVIA CRISTINA DE MENEZES SANTOS",
"RENATO MENDES FELIX",
"PETRONIO ORNELAS",
"VALMIR SALES ARAUJO",
"FERNANDA CREPALDI DE LIMA FIRMINO",
"SHARLENE SUELY DINIZ",
"ANDERSON BORGES TEIXEIRA",
"MARILENE RODRIGUES DOS SANTOS ",
"JOSE ARLINDO DA SILVA ",
"LEANDRO MASSUO NAKASHIMA",
"MARCOS ALBERTO SERAFIM",
"JEFFERSON APARECIDO ANDRIHETTI DE SOUZA",
"KLEBSOM ROGERIO DOS SANTOS ARAUJO",
"ARIVALDO VIEIRA DOS SANTOS",
"SANDRO PEREIRA LOPES JUNIOR",
"WELLINGTON ANSELMO DA CUNHA",
"ALTAMIRO MACEDO DE OLIVEIRA",
"ANDRE LUIZ XAVIER",
"RONEI MUNIZ DA CUNHA",
"NILTON SOARES CORREA FILHO",
"JAIRO BARRAS LOURENCO",
"TIAGO CANDIDO DE FREITAS",
"JOSE VITOR PEREIRA CARDOSO",
"EDNALDO DOS SANTOS FRANCISCO",
"ALESSANDRO DA CRUZ EMERENCIANO",
"ANDERSON TEIXEIRA DE FREITAS",
"LUIZ CLAUDIO BERNARDO VALENTIM ",
"VANESSA CEZAR HIGASHI DOS SANTOS",
"JOSE EDSON SARAIVA DE BARROS",
"ROGER EYVANS DOS SANTOS MARQUES SOARES ",
"RAFAEL PEREIRA DA CRUZ ",
"LUCIENE FELIX DE SOUSA",
"FRANCISCO ROSA DE ASSIS",
"MATEUS HENRIQUE VENANCIO",
"SAMUEL ANDRADE RIBEIRO",
"SERGIO CARNEIRO DE OLIVEIRA",
"FABIO ALEXANDRE SILVA SANTOS",
"JORGE APARECIDO RODRIGUES DA NOBREGA",
"EDVALDO DE FREITAS VIEIRA",
"PRISCILA SANTOS BARBOZA",
"FLAVIA REIS DA SILVA",
"WILSON OLIVEIRA FRANCO",
"BRUNO AMARAL DE SALES",
"FABIO PULGUERES DE SOUZA",
"VILMA MENDES PINTO SILVA",
"MARCELO AUGUSTO LOPES",
"EMERSON OLIVEIRA DE TOLEDO",
"FERNANDO DURANES",
"JOSE MARCOS PEREIRA DA SILVA",
"RODRIGO BICHEL DOS SANTOS",
"JUSSIENE SANTANA SANTOS",
"FELIPE GOZELOTO",
"RODNEY REZENDE LIMA CASTRO",
"JAIRO FRANCISCO PEREIRA",
"ANDRE ADAUTO CAMPOS JUNIOR",
"LIDIANE MENDES AMORIM",
"SIDNEI ROBERTO DA SILVA",
"EDSON ALVES DE HOLANDA",
"SILVERIO TORRES RAMOS",
"MARCELO RODRIGUES DE OLIVEIRA ",
"SILVIO RANGEL DA CRUZ",
"JOSE CLAUDIO DE ARAUJO MARIANO SOUZA",
"GLEICYANE APARECIDA DE SOUZA GUIMARAES",
"MICAEL MARQUES BARROS",
"FLAVIO BERNARDINO DE SENA",
"AURELIO NONATO",
"ANDRE DE CASTRO DA SILVA",
"JOICE APARECIDA ARAUJO BARRETO ",
"JESSICA CRISTINA DOS SANTOS",
"ANTONIO FRUTUOSO DA SILVA",
"ELIO FRANCISCO RIBEIRO",
"FABIANO DA SILVA CAMARGOS",
"DIEGO MARQUES PEREIRA",
"VALDIQUE OLIVEIRA DA SILVA",
"SOLANE SILVA ALMEIDA",
"EDUARDO FERREIRA LINS",
"JEFFERSON APARECIDO ANDRIHETTI DE SOUZA",
"ELIEL GONCALVES DE OLIVEIRA",
"JORGE MANOEL DE SOUZA ",
"FERNANDA VALESKA AGUIAR",
"MAYLOM HENRIQUE BUK MORAES CRUZ ",
"FABIO RODRIGUES BARBOSA",
"SAMUEL GOUVEA MAIA ",
"WAGNER DE SOUZA",
"CLEUDIA SOARES DE SALES",
"MAURO ALVES DA SILVA",
"DAYANE DA SILVA GOIS",
"TELMA REGINA OGEDA",
"MARCIO LEITE DA SILVA",
"CLAUDIA DAS DORES TEIXEIRA GIMENEZ",
"RONIE PETER LOPES DA SILVA",
"PAULO SERGIO DE OLIVEIRA",
"MAURO APARECIDO ELESBAO",
"LEANDRO JOSE TORRES",
"LUIZ CARLOS ZAMPOLLO",
"LILIAN CRISTINA FILGUEIRA DE LIMA ",
"JEFERSON ALMEIDA CAVALHEIRO",
"RAFAEL GONCALVES DA SILVA",
"MARIO SIQUEIRA",
"VANIA LUCIA DE PAIVA",
"JOILSON LUCAS ROSARIO BATISTA",
"CLAUDIO LIMA DE OLIVEIRA",
"VANESSA BARBOSA DE SOUSA SILVA",
"JOSSUE SALES DA SILVA ",
"VANDIRA DOS SANTOS ESQUIVEL CONDURU",
"MARCOS MOREIRA BARBOSA",
"CLAUDIA MARCIA NONATO ",
"JOAO LUCAS CARVALHO DE BRITO ",
"CARLOS ALBERTO ALVES SANTOS",
"VANDIRA DOS SANTOS ESQUIVEL CONDURU",
"GUILHERME HENRIQUE ALVES DA SILVA",
"JOAO VICENTE DA SILVA JUNIOR",
"ELIEL MENDES DA SILVA ASSUNCAO",
"LEILA PAULA BARROS",
"DANAEL FERREIRA",
"ADALBERTO CARLOS",
"CLAUDIO BATISTA DA SILVA ",
"FABIO SIMOR GOMES ",
"CRISTIANO RODRIGUES DE SOUZA ",
"UESCLEI OLIVEIRA BORGES",
"EVERTON ANTONIO DE CAMARGOS",
"ANTONIO SILVA CARDOSO JUNIOR",
"ALCILEIA SILVA AMARAL SANTOS",
"JOAO ELIAS DA SILVA",
"PAULO GARCIA MARTINS",
"PAULO HENRIQUE BIGAS MALAQUIAS FERREIRA",
"PAULO ROBERTO GRASSELI",
"PAULO SERGIO",
"FABRICIO LOUZADA RANGEL",
"CATIA SIMONE NOBETO",
"MARCELO CARMO OLIVEIRA",
"ADRIANO CRUZ DA SILVA",
"FERNANDO CESAR GOMES DA SILVA GALDINO DE SOUZA",
"ISAQUE DE BRITO LIMA",
"FLAVIO DOS SANTOS COTA ",
"MONICA AMALIA DOS SANTOS",
"JOSE FRANCISCO SOUSA DA SILVA",
"DARIO PEIXOTO LIMA",
"JOSE MARIO SOARES DA COSTA",
"GESSI JACI CABRAL",
"ANTONIO CARLOS CORREA GONCALVES JUNIOR",
"RICARDO RIBEIRO PESSOA",
"NERONILSON HONORATO AMORIM",
"LUCIANO NEGRAO ",
"GUILHERME CELESTINO DINIZ ",
"VALDINEI APARECIDO DE LIMA ",
"JOSE INACIO DO CARMO JUNIOR",
"WELLINGTON DE OLIVEIRA SILVERIO",
"CRISTIANO COSTA CORREA",
"THIAGO LUIZ DA SILVA BARROSO",
"FABIO SMIN",
"CLEUDIA SOARES DE SALES",
"ROBSON JOSE SPINOLA MORAIS ",
"RITA DE CASSIA DA SILVA CRUZ",
"LUCAS DE ALMEIDA PEREIRA ",
"RAFAEL DE JESUS VIEIRA",
"WAGNER HERCILIO COMIN",
"CLAUDIO PINHO BRANDAO",
"GABRIELLE SAMPAIO DOS SANTOS",
"ALEXANDRE LIMA DE SOUSA JUNIOR",
"RENATA DE OLIVEIRA",
"BRUNO BARBOSA FARNEZI",
"ROSEMERE RUFINO DIAS",
"DIEGO BITTENCOURT DA COSTA ",
"ROGERIO FERREIRA GONCALVES ",
"FRANCISCO CLEMERSON CAVALCANTE SILVA ",
"FERNANDO SOARES DE FIGUEIREDO",
"DIEGO FABIANO DA CRUZ ALVES",
"KLEBSOM ROGERIO DOS SANTOS ARAUJO",
"LAURA BARBOSA",
"ORTIAGO BENITES DE OLIVEIRA",
"JOSE CARLOS DE ALMEIDA",
"ROBSON DOS SANTOS MUNIZ ",
"FELIPE GOMES DA COSTA ",
"RODRIGO AGENOR SILVA BARBOSA",
"MARIA DA CRUZ DA COSTA SILVA",
"CARLOS OTAVIO DE ARAUJO",
"NATANAEL FAGUNDES DE LIMA ",
"JORGE APARECIDO RODRIGUES DA NOBREGA",
"RODRIGO AGENOR SILVA BARBOSA",
"DANILO PORCHAT DE ASSIS MUROLO",
"CRISTINA JUNIA SOUZA DE MATOS",
"PAULO HENRIQUE SILVA DE PAULA",
"FABIO MENDES TORRES",
"ALEXSSANDER CAMPOS TIBURCIO",
"PATRICIA DOS SANTOS",
"SERGIO LUIZ ROQUE",
"LUDILANE ARAUJO OLIVEIRA",
"JACKSON MACEDO DUTRA",
"PAULO SOUZA DOS SANTOS",
"CLOVIS FERNANDES HILARIO",
"LUIZ CARLOS DO ROCIO MACHADO ",
"CARLOS HENRIQUE CARPES ",
"PATRICIA CAMARGOS SILVA SANTOS",
"WEVERLINE NOBRE BAPTISTA",
"MAURO CESAR ISOLA",
"MAYCON BERGAMO NASCIMENTO OVERNEZ",
"CLEBER RAMOS RODRIGUES",
"SIMONE DE SOUZA ASSUMPCAO ",
"GUSTAVO COLOMBARI DA SILVA",
"LUIS CLAUDIO SIQUEIRA DE JESUS",
"WILLIAM DE PAULA FARIAS",
"ANSELMO PRADO LAMEU",
"ISAIAS LEAO DA SILVA",
"PAMELA SILVA BARBOSA",
"ELIEL NARCISO CAETANO",
"CARLOS FABIO VIANA DE LIMA",
"NOE FERREIRA DE MORAIS",
"DIOGENES VANDERLEI MIGUEL",
"LAVINA ROSA DA SILVA",
"PEDRO DE DEUS CANO MERLIN",
"HELLEN COSTA PEREIRA",
"EDUARDO DA SILVA",
"MOACY LIRA DA SILVA ",
"JOAO APARECIDO PUGLIANO ",
"JOSIAS DA COSTA BARBOSA",
"LEANDRO LACY RODRIGUES DA SILVA",
"ADRIANO PINTO BARCELOS ",
"DEJAVAN PEREIRA DOS SANTOS",
"CRISTIANE DE ARAUJO SIMAO",
"EDIVALDO FARIAS DOS SANTOS",
"ALEXANDER DA SILVA LOPES FIGUEIREDO",
"JOSE FRANCISCO BORGES DOS SANTOS",
"FERNANDO GUIZILINI DA SILVA ",
"ALEXANDRE ALVARES PEREIRA",
"MICHELE COSTA MOURA",
"CYNTHIA ALECSANDRA APARECIDA DA SILVA",
"EDNA MARIA MENDES",
"BENEDITO NAZARET TOSTA",
"MARCO ANTONIO ARAUJO",
"THIAGO LUIZ VELOSO DE OLIVEIRA",
"DOURIVAL JOSE DA SILVA",
"TEOGILA COSTA E SILVA",
"LUIZ CLAUDIO DA FONSECA DURVAL",
"LUCIANO DA COSTA DA SILVA",
"DANIEL FERNANDES ROCHA",
"FRANCISCO SANTANA BRITO DA SILVA ",
"MANOEL JOSE DOS SANTOS",
"BRUNO RODRIGUES DA SILVA ",
"WEBERT ELIZIO MACHADO",
"DIRCEU PEIXOTO PENIDO",
"LILIANE IZABEL DA SILVA",
"ANTONIO CARLOS DE ASSIS GOMES",
"ROBSON CORREA DOS SANTOS",
"RAFAEL ARAUJO PINHEIRO",
"THIAGO ANTONIO OLIVEIRA SILVA ",
"MARIO FERNANDO PENA ",
"ALEX MENDES FIDELES",
"CLEITON DE FARIA COSTA",
"JULIANA ARAUJO ALVES RODRIGUES",
"ALAN FELIX DE SOUZA",
"FABIO HENRIQUE ALVES ",
"ANDRE DE CASTRO DA SILVA",
"ANDERSON GONCALVES",
"JOSE CLAUDINO ",
"ALILIANA SILVA DE SOUSA",
"RICARDO ALEXANDRE RODES NEVES",
"EDMAR VIEIRA SOARES",
"RAMON ALVES DE SOUZA",
"ANDERSON ALEXANDRE DE ALMEIDA MAZETTO",
"CARLOS ROSEMBERG LACERDA ALEGRE DOS SANTOS",
"SILAS BALBINO",
"SUELI DANTAS SOUZA ",
"FRANCISCO DAS CHAGAS LOPES DA SILVA",
"JAIME JOSE HENRIQUES JUNIOR",
"MARCOS VINICIUS DE SOUZA ALVES",
"EVANDRO BARBOSA LEITE ",
"ROBSON ALVES DA SILVA",
"LUIZ FERNANDO SOARES RIBEIRO ",
"MARIO SANTOS",
"WANDER LIMA RODRIGUES",
"CLEIA MARCIA FERREIRA",
"MARLUCIO ROSA DE JESUS",
"ALESSANDRA MATIAS FERREIRA",
"JOAO BATISTA LEAO NETO",
"GONCALO DE AMARANTE TEIXEIRA DA SILVA",
"ALESSANDRO DE SOUZA",
"LEONARDO SOUSA PEREIRA",
"HERALDO FERREIRA DA SILVA JUNIOR",
"RONILDO ALEXANDRE DE PAULA",
"ANTONIO SANTOS DA SILVA",
"FERNANDO DA COSTA RAMOS JUNIOR",
"FABIO MONTEIRO DE ARAUJO",
"FABIO LUIZ BUZINARO",
"ELIANA CRUZ GUIMARAES",
"CLAUDIO FERREIRA DA SILVA",
"LEONARDO SOARES RIBEIRO ",
"FERNANDA VIEIRA DELGADO",
"VALDEMILSON ROSA DOS SANTOS",
"RENATO PACHECO SANTOS",
"DENILSON VICENTE DE OLIVEIRA",
"MARCOS ANTONIO DA SILVA",
"WILK ASSIS OLIVEIRA",
"JOSE RAIMUNDO FILHO",
"RENATA BENONI TEODORO",
"WILK ASSIS OLIVEIRA",
"VANESSA ACBAS MARTINELLI ",
"CATIANE OLIVEIRA ANDRADE",
"CRISTIANO DA SILVA LUIZ",
"ROGERIO SILVA SANTOS ",
"JOAO BATISTA LEAO NETO",
"CLAUDIO DA SILVA BATISTA",
"GISELLE SILVESTRE DA SILVA",
"SILVIO DE ASSIS MAMEDIO",
"BRUNO FERNANDES OLIVEIRA",
"CLAUDIO CAVALINI GUILHERME ",
"MARCELO SANTANA PRADO",
"JOAO BATISTA DE OLIVEIRA",
"VANDO ALVES DOS SANTOS",
"LUCIANA CAETANO BRAGA",
"NEY MARQUES MARTINS ",
"MISAEL PEDRO DA SILVA MARCELINO",
"ALESSANDRA APARECIDA DA SILVA ",
"CLAUDIO MARCIO DE BRITO",
"SANDRA GUERRA DOS SANTOS DEVICSI",
"CHRISTIANE GONCALVES FARIAS",
"ISAQUE MACHADO DE SOUSA",
"THIAGO ALVES DA SILVA GOMES",
"JOAO RAFAEL SOARES DOS REIS",
"AMILTON PEREIRA DE AZEREDO JUNIOR",
"ALFRANDE FERRAZ DE LIMA FILHO",
"SIDNEI ALVES DE MELLO ",
"MONICA FERNANDES ANDRADE",
"THIAGO DA SILVA COELHO",
"IZAIAS RODRIGUES BUENO",
"MAURICIO MARTINS SOUZA ",
"JOSE ARAUJO DE ALMEIDA ",
"LAISA CORREA DA SILVA",
"JOSE VILSON CARDOSO DA SILVA",
"CLAUDIA SOLIDADE SOUZA SILVA ",
"SUZANA AUGUSTA DE LOURENCO",
"ZENAIDE DE ANDRADE CHAVES",
"BRUNA LUIZA DE ALMEIDA OLIVEIRA",
"ROBSON LIMA DE SOUZA ",
"CARLOS ALBERTO LEONI ",
"ARIOVALDO ROBERTO DA SILVA",
"MARCIO VAZ DE LIMA",
"JOSE ROSA DA SILVA",
"EZEQUIAS DOMINGUES RIBEIRO ",
"DENILSON ABREU BITENCUR",
"SIMONE APARECIDA DA SILVA",
"ANDEIR CORREA BUSTAMANTE SA",
"DANIEL FAUSTINO DE SOUZA",
"FERNANDO LUCIO GOMES DE ARAUJO",
"ANA CLAUDIA DE AZEVEDO DAMASCENO",
"GETULIO ROCHA SILVA",
"NILTON CARLOS DE ALMEIDA",
"JOAQUIM FERREIRA NETO",
"MARCELA REIS SANTOS",
"MARCIO FERREIRA LINS",
"DIMAS ASSUNCAO MATEUS",
"ANDERSON VALENTIM DE OLIVEIRA",
"AMAURI ALVES D ASSUNCAO",
"JANICE TERESINHA ALVES DE CARVALHO KRAUS ",
"EMILENE RODRIGUES ",
"FELIPE DA COSTA REIS",
"KARINA TAMIKO TIAGO TSUBOI",
"DOUGLAS FERREIRA CARNEIRO",
"JUNIOR SERGIO HERMOGENES ARAUJO",
"EDIL GUILHERME DE ALMEIDA ",
"DANIELE RIBEIRO NARDON",
"ALMINA DIAS LUZ ",
"VALTER ALMEIRA DE CAMPOS",
"JORGE RODRIGUES DA SILVA",
"ANTONIO SERGIO ROTONDO",
"DAYANNE OLIVEIRA VIANA",
"WEVERSON DA SILVA SOUZA",
"CRISTIANE PIMENTA CORDEIRO DA SILVA",
"NIELSON HENRIQUE FONSECA",
"EXPEDITO RODRIGUES DA SILVA",
"CRISTIANE PIMENTA CORDEIRO DA SILVA",
"ANTONIO CARLOS DE MEDEIROS LOPES",
"BRUNO SILVA DA CONCEICAO",
"PAULO CESAR VILLALVA DA CONCEICAO",
"VALDEMIR SANTOS NOVAES",
"WESLEY NIES GONCALVES",
"LUDIMYLLA ALVES DE SOUSA",
"FABIO LOPES SAMPAIO",
"JOSE NARCISO DE LIMA",
"VINICIUS TADEU MOREIRA LEAL",
"ISAIAS VIEIRA FEITOSA",
"FERNANDO JOSE HERMENEGILDO DA SILVA",
"DIEGO RAFAEL BAZILIO DE OLIVEIRA",
"ISAIAS VIEIRA FEITOSA",
"JOSE RUI ZAGO",
"MARCELO PIRIS VARELA ",
"MICHAEL BORGES PEREIRA",
"ELIZANGELA DE SOUZA FERREIRA",
"JOSE IVAN HENRIQUE DA SILVA",
"WALTAIR FROIS DO SERRO",
"JAIRO MARTINS SILVA",
"NESTOR NUNES DE CARVALHO",
"JUSCELINA PEREIRA DE SOUZA",
"ROSICLER SOUZA DOS SANTOS",
"MARCELO BASILIO SANTANA",
"DILSON FERREIRA DA SILVA",
"JOSUE TOWNSED",
"THIAGO ANDRADE DO CARMO",
"LAERCIO MEURER",
"MARCOS PEREIRA DOS SANTOS ",
"JOSE ELISSON BARBOSA DE LIMA ",
"ANGELITA DA SILVA MAGALHAES",
"THIAGO PALOMO FIRMA",
"ELISANGELA MARIA ZAGARE RODRIGUES",
"ELISANGELA MARIA ZAGARE RODRIGUES",
"MAYCON RODRIGUES DA ROCHA",
"BENEDITO PEREIRA SOARES",
"DALCIO GOULART VIEIRA",
"JULIANA FERREIRA DE PAULA",
"CARLOS EDUARDO DE OLIVEIRA",
"FRANCISCO SANTANA BRITO DA SILVA i",
"MARCIO JOSE DOS SANTOS",
"CHRISTIANE GONCALVES FARIAS",
"GLEISON ROBERT SANTANA",
"MICHAEL MENDONCA DE SOUZA",
"ADEMIR TOME ",
"FABRICIO CAMPOS NOGUEIRA SANTOS",
"ALEXANDER VIANA BRANDAO",
"ADEMIR TOME ",
"ADEMIR TOME ",
"ROSEMERE DE SOUZA SILVA",
"JOSE RUI ZAGO",
"JULIO MARIA DE MELO",
"GEDSON JUNIOR DE SOUSA SILVA",
"PAULO SERGIO DA SILVA",
"CARLOS ANDRE DE JESUS",
"JOSE JUNHO PEREIRA GOUVEIA",
"JACI RAIMUNDO DANTAS",
"LEANDRO DA SILVA RODRIGUES ",
"ARNALDO GUERRA DE SOUZA",
"JONATAS MARTINS TAVARES ",
"LUDIMYLLA ALVES DE SOUSA",
"MAICON SANDRO VEIGA RANGEL",
            };
        }
    }
}
