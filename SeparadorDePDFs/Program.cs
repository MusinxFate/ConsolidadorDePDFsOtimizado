using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;

namespace SeparadorDePDFs;

internal static class Program
{
    private const long MaxPdfSizeBytes = (long)(10.5 * 1024 * 1024);
    private const string PartsFolderName = "PDFs partidos";
    private const string SubfoldersOption = "--subpastas";
    private const string ParallelismOption = "--paralelismo=";
    private const string PreferredGhostscriptExecutablePath =
        @"C:\Users\murilo.souza\Downloads\ghostxps-10.07.1-win64\gxpswin64.exe";

    // Preencha com os nomes das pastas cujos PDFs serão verificados.
    // Use --subpastas ou responda S no prompt para buscar esses nomes em todos os níveis.
    // Com o array vazio, a pasta informada também é verificada.
    private static readonly string[] NomesDasPastas =
    [
        // "Pasta 1",
        // "Pasta 2",
    ];

    private static readonly object LogLock = new();
    private static readonly Lazy<string?> GhostscriptExecutablePath = new(FindGhostscriptExecutable);

    static void Main(string[] args)
    {
        var includeSubfolders = false;
        var parallelism = 10;
        var pathArgs = new List<string>();
        foreach (var arg in args)
        {
            if (arg.Equals(SubfoldersOption, StringComparison.OrdinalIgnoreCase))
            {
                includeSubfolders = true;
            }
            else if (arg.StartsWith(ParallelismOption, StringComparison.OrdinalIgnoreCase))
            {
                if (!int.TryParse(arg.AsSpan(ParallelismOption.Length), out parallelism) || parallelism < 1)
                {
                    Log("Informe um inteiro positivo em --paralelismo=N.");
                    return;
                }
            }
            else
            {
                pathArgs.Add(arg);
            }
        }

        var folderPath = GetFolderPath(pathArgs);

        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            Console.WriteLine("Pasta nao encontrada.");
            return;
        }

        if (args.Length == 0)
        {
            Console.Write("Verificar subpastas em todos os níveis? (S/N): ");
            includeSubfolders = string.Equals(Console.ReadLine()?.Trim(), "s", StringComparison.OrdinalIgnoreCase);
        }

        var pdfFiles = GetFoldersToScan(folderPath, includeSubfolders)
            .SelectMany(folder => folder.EnumerateFiles("*.pdf", SearchOption.TopDirectoryOnly))
            .Where(file => !IsGeneratedPart(file.Name) && file.Length > 0)
            .OrderBy(file => file.FullName)
            .ToArray();

        if (pdfFiles.Length == 0)
        {
            Log("Nenhum PDF encontrado nas pastas selecionadas.");
            return;
        }

        Log($"PDFs encontrados: {pdfFiles.Length}");
        Log($"Limite por parte: {FormatSize(MaxPdfSizeBytes)}");

        if (parallelism == 1)
        {
            foreach (var pdfFile in pdfFiles)
            {
                SplitPdfIfNeeded(pdfFile);
            }
        }
        else
        {
            Parallel.ForEach(pdfFiles, new ParallelOptions { MaxDegreeOfParallelism = parallelism }, SplitPdfIfNeeded);
        }

        Log("Finalizado.");
    }

    private static IEnumerable<DirectoryInfo> GetFoldersToScan(string folderPath, bool includeSubfolders)
    {
        var root = new DirectoryInfo(folderPath);
        var selectedNames = new HashSet<string>(NomesDasPastas, StringComparer.OrdinalIgnoreCase);
        if (selectedNames.Count == 0)
        {
            yield return root;
            if (!includeSubfolders)
            {
                yield break;
            }
        }

        var enumerationOptions = new EnumerationOptions
        {
            AttributesToSkip = FileAttributes.ReparsePoint,
            IgnoreInaccessible = false
        };
        var pendingFolders = new Stack<DirectoryInfo>();
        pendingFolders.Push(root);

        while (pendingFolders.TryPop(out var folder))
        {
            foreach (var subfolder in folder.EnumerateDirectories("*", enumerationOptions))
            {
                if (subfolder.Name.Equals(PartsFolderName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (selectedNames.Count == 0 || selectedNames.Contains(subfolder.Name))
                {
                    yield return subfolder;
                }

                if (includeSubfolders)
                {
                    pendingFolders.Push(subfolder);
                }
            }
        }
    }

    private static string? GetFolderPath(List<string> args)
    {
        if (args.Count > 0)
        {
            return string.Join(' ', args).Trim('"');
        }

        Console.Write("Informe o caminho da pasta onde estão os PDFs: ");
        return Console.ReadLine()?.Trim().Trim('"');
    }

    private static void SplitPdfIfNeeded(FileInfo pdfFile)
    {
        var pdfPath = pdfFile.FullName;
        try
        {
            var originalSize = pdfFile.Length;

            if (originalSize <= MaxPdfSizeBytes)
            {
                Log($"[OK] {Path.GetFileName(pdfPath)} já está dentro do limite ({FormatSize(originalSize)}).");
                return;
            }

            ProcessOversizedPdf(pdfFile);
        }
        catch (Exception ex)
        {
            Log($"[ERRO] {pdfPath} => {ex.Message}");
        }
    }

    private static void ProcessOversizedPdf(FileInfo pdfFile)
    {
        var workspace = Directory.CreateTempSubdirectory("separador-pdf-");
        try
        {
            var sourcePath = pdfFile.FullName;
            var root = Path.GetPathRoot(sourcePath)!;
            if (root.StartsWith(@"\\", StringComparison.Ordinal) || new DriveInfo(root).DriveType == DriveType.Network)
            {
                sourcePath = Path.Combine(workspace.FullName, pdfFile.Name);
                File.Copy(pdfFile.FullName, sourcePath);
            }

            using var input = OpenWithRetry(sourcePath);
            var pageCount = input.PageCount;
            if (pageCount == 0)
            {
                Log($"[VAZIO] {pdfFile.FullName}");
                return;
            }

            var workingPartsFolder = workspace.FullName;
            var partNumber = 1;
            var createdParts = new List<string>();
            var startPageIndex = 0;

            while (startPageIndex < pageCount)
            {
                var endPageIndex = FindLargestPageRangeEnd(input, sourcePath, startPageIndex, pageCount - 1);

                if (endPageIndex >= startPageIndex)
                {
                    SaveRangeOrSplitFailures(
                        input,
                        sourcePath,
                        workingPartsFolder,
                        startPageIndex,
                        endPageIndex,
                        ref partNumber,
                        createdParts);
                    startPageIndex = endPageIndex + 1;
                    continue;
                }

                ExtractPagesSeparately(
                    input,
                    sourcePath,
                    workingPartsFolder,
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

            var oversizedCount = oversizedParts.Count(part => GetFileSize(part) > MaxPdfSizeBytes);

            if (createdParts.Count == 0)
            {
                Log($"[ERRO] {pdfFile.Name} => nenhuma parte foi criada.");
                return;
            }

            var partsFolderPath = GetPartsFolderPath(pdfFile.FullName);
            PublishParts(pdfFile.FullName, partsFolderPath, createdParts);

            if (oversizedCount > 0)
            {
                Log($"[ERRO LIMITE] {pdfFile.Name} gerou {oversizedCount} parte(s) ainda acima de {FormatSize(MaxPdfSizeBytes)}.");
                return;
            }

            Log($"[DIVIDIDO] {pdfFile.Name} => {createdParts.Count} parte(s) em {partsFolderPath}. Original mantido.");
        }
        finally
        {
            try
            {
                workspace.Delete(recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Log($"[AVISO LIMPEZA] {workspace.FullName} => {ex.Message}");
            }
        }
    }

    private static void SaveRangeOrSplitFailures(
        PdfDocument input,
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
            SavePageRange(input, startPageIndex, endPageIndex, partPath);
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
                input,
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
            input,
            pdfPath,
            partsFolderPath,
            startPageIndex,
            midPageIndex,
            ref partNumber,
            createdParts);

        SaveRangeOrSplitFailures(
            input,
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
            Log($"[COMPRESS SKIP] Ghostscript não encontrado. Arquivo: {pdfPath}");
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
        PdfDocument input,
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
                SavePageRange(input, pageIndex, pageIndex, singlePagePath);
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
            Log($"[GS AVISO] Usando {PreferredGhostscriptExecutablePath}, mas esse executável é GhostXPS e pode não ler PDF de entrada.");
            return PreferredGhostscriptExecutablePath;
        }

        return null;
    }

    private static int FindLargestPageRangeEnd(PdfDocument input, string sourcePdfPath, int startPageIndex, int lastPageIndex)
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
                using var stream = new MemoryStream();
                using var output = CreatePdfDocument();
                AddPages(input, startPageIndex, mid, output);
                output.Save(stream, false);
                candidateSize = stream.Length;
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

    private static void SavePageRange(PdfDocument input, int startPageIndex, int endPageIndex, string outputPath)
    {
        using var output = CreatePdfDocument();
        AddPages(input, startPageIndex, endPageIndex, output);
        output.Save(outputPath);
    }

    private static void AddPages(PdfDocument input, int startPageIndex, int endPageIndex, PdfDocument output)
    {
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
            throw new InvalidOperationException("PDF extraído não foi criado ou ficou com 0 bytes.");
        }

        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);

        if (document.PageCount != expectedPageCount)
        {
            throw new InvalidOperationException($"PDF extraído ficou com {document.PageCount} página(s), esperado {expectedPageCount}.");
        }

        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var page = document.Pages[pageIndex];

            if (page.Contents.Elements.Count == 0)
            {
                throw new InvalidOperationException($"Página {pageIndex + 1} do PDF extraído ficou sem stream de conteúdo.");
            }
        }
    }

    private static string GetPartsFolderPath(string pdfPath)
    {
        var directory = Path.GetDirectoryName(pdfPath) ?? string.Empty;
        return Path.Combine(directory, PartsFolderName);
    }

    private static void PublishParts(string pdfPath, string partsFolderPath, List<string> createdParts)
    {
        Directory.CreateDirectory(partsFolderPath);
        var pendingParts = new List<(string TemporaryPath, string FinalPath)>(createdParts.Count);
        var publishedCount = 0;

        try
        {
            foreach (var partPath in createdParts)
            {
                var temporaryPath = Path.Combine(partsFolderPath, $".{Guid.NewGuid():N}.tmp");
                var finalPath = Path.Combine(partsFolderPath, Path.GetFileName(partPath));
                pendingParts.Add((temporaryPath, finalPath));
                File.Copy(partPath, temporaryPath);
            }

            foreach (var (temporaryPath, finalPath) in pendingParts)
            {
                File.Move(temporaryPath, finalPath, overwrite: true);
                publishedCount++;
            }

            var publishedPaths = pendingParts.Select(part => part.FinalPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var pattern = $"{Path.GetFileNameWithoutExtension(pdfPath)} - parte *.pdf";
            foreach (var existingPart in Directory.EnumerateFiles(partsFolderPath, pattern, SearchOption.TopDirectoryOnly))
            {
                if (!publishedPaths.Contains(existingPart))
                {
                    File.Delete(existingPart);
                }
            }
        }
        finally
        {
            foreach (var (temporaryPath, _) in pendingParts.Skip(publishedCount))
            {
                try
                {
                    File.Delete(temporaryPath);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    Log($"[AVISO LIMPEZA] {temporaryPath} => {ex.Message}");
                }
            }
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
