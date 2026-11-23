using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

[assembly: SupportedOSPlatform("windows")]

namespace ConsolidacaoDePDFs
{
    internal class Program
    {
        private static readonly string LogFolderPath = @"C:\Users\murilo.souza\Desktop\Testes - Decadencia\Logs\";
        private static readonly string LogPath = $@"{LogFolderPath}consolidacao_log_infos.txt";

        private static readonly string LogNaoEncontrouPastaPath =
            $@"{LogFolderPath}consolidacao_log_existe_excel_nao_existe_pasta.txt";

        private static readonly string LogNaoEncontradoExcelPath =
            $@"{LogFolderPath}consolidacao_log_existe_pasta_nao_existe_no_excel.txt";

        private static readonly string LogPastasVaziasPath = $@"{LogFolderPath}consolidacao_log_pastas_vazias.txt";
        private static readonly string LogJaExisteConsolidacao = $@"{LogFolderPath}consolidacao_log_ja_consolidado.txt";
        private static readonly object _lock = new object();
        private static readonly object _wordLock = new object();
        private const long MaxConsolidatedPdfSizeBytes = (long)(14.5 * 1024 * 1024);

        private const string GhostscriptMissingLogMessage =
            "Ghostscript não encontrado. Não é possível reduzir uma página única acima de 14,5 MB apenas com split.";

        private static readonly Lazy<string?> GhostscriptExecutablePath = new(FindGhostscriptExecutable);

        private static readonly string BasePath =
        //@"\\SRV-FS02-BBM\Empresas$\Empresas\AMCOR\1. CONTROLE\3. CP NAS RTS- DECADÊNCIA\3. BOOKS";
        //@"C:\Users\murilo.souza\Desktop\Testes - Decadencia\Fazer";
        @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\13. PAFs\19613.7393632022-72 - 6ª COMPENSAÇÃO 07.2022\BOOK";

        static void Main(string[] args)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var dividirPdfsPorTamanho = DeveDividirPdfsPorTamanho();

            Log(dividirPdfsPorTamanho
                ? "Divisão dos PDFs por tamanho: SIM"
                : "Divisão dos PDFs por tamanho: NÃO (somente consolidação)");

            var nomes = GetNomes();

            var nomesSet = new HashSet<string>(nomes);

            var folders = Directory.GetDirectories(BasePath);

            var matchingFolders = new List<string>();
            var encontrados = new HashSet<string>();

            var nomesPastas = new List<string>();

            foreach (var folder in folders)
            {
                var folderName = Path.GetFileName(folder);
                nomesPastas.Add(folderName);

                if (nomesSet.Contains(folderName))
                {
                    matchingFolders.Add(folder);
                    encontrados.Add(folderName);
                }
            }

            var naoEncontrados = nomesSet
                .Where(n => !encontrados.Contains(n))
                .ToList();

            File.WriteAllLines(LogNaoEncontrouPastaPath, naoEncontrados);

            var pastasNaoNaLista = nomesPastas
                .Where(nome => !nomesSet.Contains(nome))
                .ToList();

            File.WriteAllLines(LogNaoEncontradoExcelPath, pastasNaoNaLista);

            Log($"Pastas para processar: {matchingFolders.Count}");
            Log($"Não encontrados: {naoEncontrados.Count}");
            Log($"Pastas fora da lista: {pastasNaoNaLista.Count}");

            var options = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach(
                matchingFolders,
                options,
                folder => ProcessFolder(folder, dividirPdfsPorTamanho));

            sw.Stop();
            Log($"FINALIZADO em {sw.Elapsed}");
        }

        static bool DeveDividirPdfsPorTamanho()
        {
            while (true)
            {
                Console.Write("Deseja dividir os PDFs consolidados com base no tamanho? (S/N): ");
                var resposta = Console.ReadLine()?.Trim();

                if (resposta?.Equals("S", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return true;
                }

                if (resposta?.Equals("N", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return false;
                }

                Console.WriteLine("Opção inválida. Digite S para sim ou N para não.");
            }
        }

        static void ProcessFolder(string folder, bool dividirPdfsPorTamanho)
        {
            try
            {
                var folderName = Path.GetFileName(folder);
                var outputPath = Path.Combine(folder, $"RT {folderName}.pdf");

                if (File.Exists(outputPath))
                {
                    if (dividirPdfsPorTamanho)
                    {
                        SplitPdfIfNeeded(outputPath);
                    }

                    Log($"[SKIP] {folderName}");
                    LogJaExisteConsolidacao_($"{folderName}");
                    return;
                }

                if (HasConsolidatedParts(folder, folderName))
                {
                    if (dividirPdfsPorTamanho)
                    {
                        SplitExistingConsolidatedParts(folder, folderName);
                    }

                    Log($"[SKIP PARTS] {folderName}");
                    LogJaExisteConsolidacao_($"{folderName}");
                    return;
                }

                ConvertDocFilesToPdf(folder);

                var files = Directory
                    .EnumerateFiles(folder, "*.pdf", SearchOption.TopDirectoryOnly)
                    .Where(HasContent)
                    .Where(f => !IsConsolidatedPdf(f, folderName))
                    .OrderBy(f => f)
                    .ToArray();

                if (files.Length == 0)
                {
                    Log($"[EMPTY] {folderName}");
                    LogPastasVazias(folderName);
                    return;
                }

                using var outPdf = new PdfDocument();
                ConfigurePdfOptions(outPdf);

                foreach (var file in files)
                {
                    TryAppendPdf(file, outPdf);
                }

                if (outPdf.PageCount > 0)
                {
                    outPdf.Save(outputPath);

                    if (dividirPdfsPorTamanho)
                    {
                        SplitPdfIfNeeded(outputPath);
                    }

                    Log($"[OK] {folderName}");
                }
                else
                {
                    Log($"[NO PAGES] {folderName}");
                }
            }
            catch (Exception ex)
            {
                Log($"[ERRO PASTA] {folder} => {ex.Message}");
            }
        }

        static void ConvertDocFilesToPdf(string folder)
        {
            var docFiles = Directory
                .EnumerateFiles(folder, "*.doc", SearchOption.TopDirectoryOnly)
                .Concat(Directory.EnumerateFiles(folder, "*.docx", SearchOption.TopDirectoryOnly))
                .Where(IsWordDocument)
                .Where(HasContent)
                .Where(f => !File.Exists(Path.ChangeExtension(f, ".pdf")))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(f => f)
                .ToArray();

            if (docFiles.Length == 0)
            {
                return;
            }

            try
            {
                lock (_wordLock)
                {
                    RunOnStaThread(() => ConvertDocumentsToPdfWithWord(docFiles));
                }
            }
            catch (Exception ex)
            {
                Log($"[ERRO WORD] {folder} => {ex.Message}");
            }
        }

        static bool IsWordDocument(string file)
        {
            var extension = Path.GetExtension(file);
            return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
                   || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase);
        }

        static bool HasContent(string file)
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

        static void ConvertDocumentsToPdfWithWord(string[] docFiles)
        {
            object? wordApplication = null;

            try
            {
                var wordType = Type.GetTypeFromProgID("Word.Application")
                               ?? throw new InvalidOperationException(
                                   "Microsoft Word não está instalado ou não está registrado para automação COM.");

                wordApplication = Activator.CreateInstance(wordType)
                                  ?? throw new InvalidOperationException("Não foi possível iniciar o Microsoft Word.");

                dynamic word = wordApplication;
                word.Visible = false;
                word.DisplayAlerts = 0;

                foreach (var docPath in docFiles)
                {
                    object? document = null;
                    var pdfPath = Path.ChangeExtension(docPath, ".pdf");

                    try
                    {
                        document = word.Documents.Open(
                            FileName: docPath,
                            ConfirmConversions: false,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Visible: false);

                        dynamic doc = document;
                        doc.SaveAs2(FileName: pdfPath, FileFormat: 17);
                        doc.Close(SaveChanges: false);
                        document = null;

                        Log($"[DOC->PDF] {docPath}");
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERRO DOC] {docPath} => {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(document);
                    }
                }

                word.Quit(SaveChanges: false);
                wordApplication = null;
            }
            finally
            {
                ReleaseComObject(wordApplication);
            }
        }

        static void RunOnStaThread(Action action)
        {
            Exception? exception = null;

            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    exception = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (exception is not null)
            {
                throw exception;
            }
        }

        static void SplitPdfIfNeeded(string pdfPath)
        {
            if (!File.Exists(pdfPath))
            {
                return;
            }

            if (GetFileSize(pdfPath) <= MaxConsolidatedPdfSizeBytes)
            {
                return;
            }

            int pageCount;
            using (var input = OpenWithRetry(pdfPath))
            {
                pageCount = input.PageCount;
            }

            var partNumber = 1;
            var createdParts = new List<string>();
            var startPageIndex = 0;

            while (startPageIndex < pageCount)
            {
                var endPageIndex = FindLargestPageRangeEnd(pdfPath, startPageIndex, pageCount - 1);

                if (endPageIndex >= startPageIndex)
                {
                    var partPath = GetPartPath(pdfPath, partNumber++);
                    SavePageRange(pdfPath, startPageIndex, endPageIndex, partPath);
                    createdParts.Add(partPath);
                    startPageIndex = endPageIndex + 1;
                    continue;
                }

                var oversizedSinglePagePath = GetPartPath(pdfPath, partNumber++);
                SavePageRange(pdfPath, startPageIndex, startPageIndex, oversizedSinglePagePath);
                TryCompressPdfToLimit(oversizedSinglePagePath);
                createdParts.Add(oversizedSinglePagePath);
                Log(
                    $"[SPLIT WARN] Página {startPageIndex + 1} sozinha ficou acima de 14,5 MB: {oversizedSinglePagePath} ({FormatSize(GetFileSize(oversizedSinglePagePath))})");
                startPageIndex++;
            }

            var oversizedParts = createdParts
                .Where(part => GetFileSize(part) > MaxConsolidatedPdfSizeBytes)
                .ToArray();

            foreach (var oversizedPart in oversizedParts)
            {
                TryCompressPdfToLimit(oversizedPart);
            }

            oversizedParts = createdParts
                .Where(part => GetFileSize(part) > MaxConsolidatedPdfSizeBytes)
                .ToArray();

            if (oversizedParts.Length == 0)
            {
                File.Delete(pdfPath);
                Log($"[SPLIT] {pdfPath} => {createdParts.Count} partes");
                return;
            }

            Log(
                $"[SPLIT ERRO LIMITE] {pdfPath} gerou {oversizedParts.Length} parte(s) ainda acima de 14,5 MB. Esses casos exigem recompressão/redução da página, não apenas split.");
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

        static int FindLargestPageRangeEnd(string sourcePdfPath, int startPageIndex, int lastPageIndex)
        {
            var low = startPageIndex;
            var high = lastPageIndex;
            var best = startPageIndex - 1;

            while (low <= high)
            {
                var mid = low + ((high - low) / 2);
                var candidateSize = GetPageRangePdfSize(sourcePdfPath, startPageIndex, mid);

                if (candidateSize <= MaxConsolidatedPdfSizeBytes)
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

        static long GetPageRangePdfSize(string sourcePdfPath, int startPageIndex, int endPageIndex)
        {
            using var stream = new MemoryStream();
            SavePageRange(sourcePdfPath, startPageIndex, endPageIndex, stream);
            return stream.Length;
        }

        static void SavePageRange(string sourcePdfPath, int startPageIndex, int endPageIndex, string outputPath)
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            using var output = CreatePdfDocument();
            AddPages(sourcePdfPath, startPageIndex, endPageIndex, output);
            output.Save(outputPath);
        }

        static void SavePageRange(string sourcePdfPath, int startPageIndex, int endPageIndex, Stream outputStream)
        {
            using var output = CreatePdfDocument();
            AddPages(sourcePdfPath, startPageIndex, endPageIndex, output);
            output.Save(outputStream, false);
        }

        static void AddPages(string sourcePdfPath, int startPageIndex, int endPageIndex, PdfDocument output)
        {
            using var input = OpenWithRetry(sourcePdfPath);

            for (var pageIndex = startPageIndex; pageIndex <= endPageIndex; pageIndex++)
            {
                output.AddPage(input.Pages[pageIndex]);
            }
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

        static string GetPartPath(string pdfPath, int partNumber)
        {
            var directory = Path.GetDirectoryName(pdfPath) ?? string.Empty;
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfPath);
            return Path.Combine(directory, $"{fileNameWithoutExtension} - PARTE {partNumber:00}.pdf");
        }

        static bool HasConsolidatedParts(string folder, string folderName)
        {
            return Directory
                .EnumerateFiles(folder, $"RT {folderName} PARTE *.pdf", SearchOption.TopDirectoryOnly)
                .Any();
        }

        static void SplitExistingConsolidatedParts(string folder, string folderName)
        {
            var consolidatedParts = Directory
                .EnumerateFiles(folder, $"RT {folderName} PARTE *.pdf", SearchOption.TopDirectoryOnly)
                .OrderBy(f => f)
                .ToArray();

            foreach (var consolidatedPart in consolidatedParts)
            {
                SplitPdfIfNeeded(consolidatedPart);
            }
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

        static void TryAppendPdf(string file, PdfDocument output)
        {
            try
            {
                using var input = OpenWithRetry(file);
                CopyPages(input, output);
            }
            catch (Exception ex)
            {
                Log($"[ERRO ARQUIVO] {file} => {ex.Message}");
            }
        }

        static PdfDocument OpenWithRetry(string file, int maxRetries = 3)
        {
            for (int i = 0; i < maxRetries; i++)
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

            throw new Exception($"Falha ao abrir PDF: {file}");
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
                File.AppendAllText(LogPath, linha + Environment.NewLine);
            }
        }

        static void LogPastasVazias(string msg)
        {
            var linha = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                File.AppendAllText(LogPastasVaziasPath, linha + Environment.NewLine);
            }
        }

        static void LogJaExisteConsolidacao_(string msg)
        {
            var linha = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                File.AppendAllText(LogJaExisteConsolidacao, linha + Environment.NewLine);
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
