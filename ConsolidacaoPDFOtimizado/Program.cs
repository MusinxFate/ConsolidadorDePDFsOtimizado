using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ConsolidacaoDePDFs
{
    internal class Program
    {
        private static readonly string LogFolderPath = @"C:\Users\murilo.souza\Desktop\Logs - Henrique\";
        private static readonly string LogPath = $@"{LogFolderPath}consolidacao_log_infos.txt";
        private static readonly string LogNaoEncontrouPastaPath = $@"{LogFolderPath}consolidacao_log_existe_excel_nao_existe_pasta.txt";
        private static readonly string LogNaoEncontradoExcelPath = $@"{LogFolderPath}consolidacao_log_existe_pasta_nao_existe_no_excel.txt";
        private static readonly string LogPastasVaziasPath = $@"{LogFolderPath}consolidacao_log_pastas_vazias.txt";
        private static readonly string LogJaExisteConsolidacao = $@"{LogFolderPath}consolidacao_log_ja_consolidado.txt";
        private static readonly object _lock = new object();

        private static readonly string BasePath =
        //@"\\SRV-FS02-BBM\Empresas$\Empresas\AMCOR\1. CONTROLE\3. CP NAS RTS- DECADÊNCIA\3. BOOKS";
        @"C:\Users\murilo.souza\Desktop\Teste - Modo Deus";

        static void Main(string[] args)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

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
                MaxDegreeOfParallelism = Math.Min(Environment.ProcessorCount, -1)
            };

            Parallel.ForEach(matchingFolders, options, ProcessFolder);

            sw.Stop();
            Log($"FINALIZADO em {sw.Elapsed}");
        }

        static void ProcessFolder(string folder)
        {
            try
            {
                var folderName = Path.GetFileName(folder);
                var outputPath = Path.Combine(folder, $"{folderName} - CONSOLIDADO.pdf");

                if (File.Exists(outputPath))
                {
                    Log($"[SKIP] {folderName}");
                    LogJaExisteConsolidacao_($"{folderName}");
                    return;
                }

                var pdfFiles = Directory
                    .EnumerateFiles(folder, "*.pdf", SearchOption.TopDirectoryOnly)
                    .OrderBy(f => f)
                    .ToArray();

                if (pdfFiles.Length == 0)
                {
                    Log($"[EMPTY] {folderName}");
                    LogPastasVazias(folderName);
                    return;
                }

                using var outPdf = new PdfDocument();

                foreach (var file in pdfFiles)
                {
                    TryAppendPdf(file, outPdf);
                }

                if (outPdf.PageCount > 0)
                {
                    outPdf.Save(outputPath);
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
                    "JOSE BARBOSA SOBRINHO",
                    "CARLOS ANTONIO ROCHA",
                    "CLARECI LEITE DE SENA",
                    "EDVALDO GOMES",
                    "FABIO JUNIOR GIMENES CASTRO",
                    "CICERO DE OLIVEIRA",
                    "ORLANDO ALVES DA SILVA",
                    "JOSE DO BONFIM AMORIM LOPES",
                    "GENECI FERREIRA DE OLIVEIRA",
                    "WILLIAN DUARTE DE LIMA",
                    "FERNANDO CAETANO DOS SANTOS",
                    "RAIMUNDO GONÇALVES  DA SILVA NETO",
                    "AILTON BARBOSA DE OLIVEIRA",
                    "PAULO CESAR NEGRÃO",
                    "ANDERSON FELIX DE CARVALHO",
                    "FERNANDO PIMENTA",
                    "PAULO CESAR MORAES",
                    "ISRAEL DO CARMO RIBEIRO",
                    "VINICIUS ALONSO RODRIGUES",
                    "LORIVALDO DE OLIVEIRA LIMA",
                    "NODIR LOPES DE LIMA",
                    "LEANDRO SANTOS DA SILVA",
                    "ANISIO DE ALENCAR",
                    "JHEINISSON SANTOS SILVA",
                    "JOSE ROBERTO GRACIANO",
                    "RONALDO GONÇALVES VIEIRA",
                    "MANUEL DOS SANTOS RAMOS",
                    "PAULO HENRIQUE ARAUJO DE OLIVEIRA",
                    "NELSON NOVAIS DE SOUZA",
                    "LAERTE APARECIDO DA COSTA",
                    "EDMARK ABRANGES MARTINS",
                    "VANDERLEI RODRIGUES DA SILVA",
                    "CLEYTON DE ANDRADE DA COSTA",
                    "WAGNER APARECIDO BRANDETTI",
                    "LUIZ CARLOS DE MELLO CASEMIRO",
                    "EDUARDO TADEU CASSILHAS",
                    "RAFAEL DA SILVA",
                    "ANTONIO LICHA NETO",
                    "EDUARDO PEREIRA DE MAGALHÃES",
                    "MARCIO DA ROSA CARDOSO",
                    "RODERNEI DE OLIVEIRA TAVARES",
                    "DJALMA FREITAS IGNÁCIO",
                    "EDER HENRIQUE NUNES DA SILVA",
                    "WAGNER CANELLO",
                    "ELVIS ARAUJO",
                    "LUIZ FERNANDES DA SILVA",
                    "ANDRE FELIPE ALVES DA SILVA",
                    "ERIVAN MIRANDA BASTOS",
                    "SERGIO LAZARO DA CRUZ",
                    "ROBSON GONCALVES MATOS",
                    "SEBASTIAO BORRERO DE PADUA JUNIOR",
                    "ENRIQUE DIAS DE OLIVEIRA",
                    "ANTONIO DA SILVA FERNANDES",
                    "ANTONIO FRANCISCO DOS SANTOS OLIVEIRA",
                    "AGUINALDO DE SOUZA",
                    "JULIO CESAR BARBOSA",
                    "DAVID APARECIDO MARRICHI",
                    "JOSE EDVALDO ALVES DE MORAIS",
                    "REINALDO ZOE DE MELO",
                    "THIAGO RODRIGUES",
                    "JOAO CARLOS DA PIEDADE",
                    "CLAUDEMIR APARECIDO DE SOUZA",
                    "FLAVIO HENRIQUE DE SOUZA OLIVEIRA",
                    "PAULO RODRIGO LUCAS",
                    "CLENILTON DE MOURA MARTIRIOS",
                    "PEDRO CARDOSO DA SILVA",
                    "JEFERSON TEIXEIRA BATISTA DUTRA",
                    "REGINALDO ALVES COELHO",
                    "RAFAEL CELIS POLICARPO",
                    "PAULO ANTONIO DE OLIVEIRA",
                    "EDSON ALVES DE SOUZA",
                    "MARCOS DE OLIVEIRA LIMA",
                    "GILBERTO DE SOUZA",
                    "ALEX SANDER RUBERTI",
                    "ALEXANDRO RODRIGUES COSTA",
                    "TIAGO FERNANDO LINO",
                    "LUIS CARLOS FERNANDES",
                    "ROBSON LUIZ GALDINO",
                    "IGOR DA SILVA",
                    "ROGERIO MARINO SALES DOS SANTOS",
                    "MARCIO MEDEIROS",
                    "WILLIAN SANTOS DA COSTA",
                    "ANDERSON FABIANO CAIRES LIMA",
                    "ROBSON MODESTO PARENTE",
                    "GILLIARD MESSIAS SIMONCINI",
                    "ANILTON JOSE DE SOUZA",
                    "SILAS VERAS BATISTA DE SOUZA",
                    "gabrigabriel",
                    "GABRIEL GONCALVES DE PAIVA",
                    "NATANAEL NASCIMENTO DA SILVA",
                    "ANTONIO CARLOS DA SILVA",
                    "MOACYR NABOR LEONEL NETO",
                    "MARCIO ANTONIO CORDEIRO",
                    "EDMILSON ANTONIO DA SILVA",
                    "JOSE MARIA CESAR DA LUZ",
                    "CLAUDEMIR STEINLE DE SOUZA",
                    "PAULO APARECIDO MORENO",
                    "ISABEL MARIA DA SILVA",
                    "IVAN SILVA DANTAS",
                    "LUCIMAR ESPEDITO MARQUES",
                    "FERNANDO GANGINI",
                    "RODRIGO ANTONIO FERRAZ",
                    "JHONATAN ELIAS RAMOS",
                    "MAURICIO PEREIRA DA SILVA",
                    "RODRIGO DA SILVA",
                    "EDUARDO ALVES RODRIGUES",
                    "JAIR DE SOUZA LIMA",
                    "FERNANDO CABRAL",
                    "IVANILDO VERSURI",
                    "ROBSON POMPEU DOS SANTOS",
                    "FABIO JOSE VIANA",
                    "WELLINGTON GUIMARAES",
                    "CLAUDEMIR MESSIAS DA SILVA",
                    "JOSE HELENO BEZERRA DE SOUZA",
                    "CLAUDECI PAES BATISTA",
                    "JOÃO AILTON DOS SANTOS",
                    "LEANDRO PACHECO AZEVEDO",
                    "JOSEMAR DE SOUZA LUZ",
                    "GIOVANI ALVES DA VEIGA",
                    "MOISES TEIXEIRA SILVA",
                    "MARCELO TADEU DE MORAES",
                    "CLEBERSON CUSTODIO",
                    "MARIA VILANI ALVES DE QUEIROZ",
                    "JOSE ZENALDO DE OLIVEIRA",
                    "MIGUEL SCHMOELLER RODRIGUES",
                    "LENILSON DA SILVA BEZERRA",
                    "ISAC DOS SANTOS IZABEL",
                    "GORDIANO FERNANDES DA SILVA",
                    "CARLOS JOSE DA SILVA",
                    "JOSE SIMAO DA SILVA",
                    "ANDRE DE FREITAS SCABORA",
                    "WELLINGTON BATISTA GARCIA",
                    "WELINTON PEREIRA DE LACERDA",
                    "NADIR CARVALHO BARBOSA NEVES",
                    "ALISSON ALVES COELHO",
                    "DIOGO MARTINS SOBREIRA",
                    "MARCIO AGNALDO SANHES",
                    "RAIMUNDO CARLOS ROCHA DE SOUSA",
                    "MARCIO EMERSON MARIANO DA SILVA",
                    "DENIS VIEIRA FERNANDES",
                    "AMERICO ANTONIO DA SILVA",
                    "RONALDO HITOSHI MURAKAMI",
                    "PEDRO FELIPE ALVES RAMOS",
                    "ANDERSON APARECIDO DO NASCIMENTO RODRIGUES",
                    "EDMAR DE SOUZA RAYMUNDO",
                    "WELTON RAMOS DE LIMA",
                    "JEFFERSON DE FARIAS GOMES",
                    "VALTER COSTA",
                    "RENAN EMERENCIANO INACIO",
                    "ANDRE SANTOS DE OLIVEIRA",
                    "HENRIQUE SIQUEIRA DA SILVA",
                    "LEANDRO SANTOS DA SILVA",
                    "MARCELO CAVALHEIRO",
                    "ROBSON ALVES DA SILVA",
                    "NATANAEL SOARES DE OLIVEIRA",
                    "MAIKE TEODORO DOS SANTOS",
                    "SERGIO CORREA",
                    "VALCIRO FERNANDES DA CRUZ",
                    "REGIS VIEIRA BARBOSA",
                    "JUSCINEI SANTOS DE ALMEIDA",
                    "FERNANDO MORANDI",
                    "THIAGO MARTINS MIGUEL",
                    "REGINALDO MELO DA SILVA",
                    "GILBERTO JOSE DOS SANTOS",
                    "LUCAS EDUARDO DE OLIVEIRA",
                    "JULIO CESAR MENDONCA DE ARRUDA",
                    "CAIO LIMA DE MACEDO",
                    "CELIO APARECIDO PEREIRA",
                    "LEONILDO SANTOS DE SOUSA",
                    "ROGERSON DE PAULA CASTRO",
                    "MARCELO DA SILVA",
                    "SILVANO FERREIRA DA SILVA",
                    "SANDRO LAURENTINO DE SOUZA",
                    "MARCELO GOMES DA SILVA JUNIOR",
                    "ISAIAS NERES DE SOUZA",
                    "WANDERLEI ROCHA GERALDO",
                    "MARCIO VIEIRA MARTINS",
                    "LUCAS FREITAS PINCETTA",
                    "MARCOS AUGUSTO FAUST CARIOCA RAMOS",
                    "EUGENIO PINHEIRO FILHO",
                    "REINALDO RIBEIRO",
                    "DOUGLAS PEREIRA OLIVEIRA",
                    "CARLOS EDUARDO DE CASTRO",
                    "BRUNO HENRIQUE LOIOLA",
                    "WILLINGTON DO ESPIRITO SANTO GOMES",
                    "CARLOS PESENTI",
                    "WAGNER FELIPE SERNICHIARI",
                    "JOÃO ANDRÉ VIEIRA LOPES",
                    "WESLER ANDERSON BIAZIM",
                    "CIRINEU ADRIANO",
                    "REINALDO TADEU RODRIGUES",
                    "SEBASTIÃO DOS SANTOS FERREIRA",
                    "HUMBERTO DANTAS DOS SANTOS",
                    "ALEXSANDRO BARBOSA DA SILVA",
                    "ALEX DE BRITO LACERDA",
                    "REINALDO GUILHERME DE AGUIAR",
                    "MOACIR URBANO",
                    "VANDERLEI APARECIDO PINHEIRO",
                    "ANDERSON CORREA QUARESMA",
                    "JEFFERSON JONATHAS DE FREITAS",
                    "JAIR GONCALVES DE JESUS",
                    "TATIANE BEZERRA DO NASCIMENTO CARVALHO",
                    "FERNANDO SILVA FLYGARE RAZO",
                    "MARCIO ROBERTO RAYMUNDO",
                    "EDUARDO DIAS DE MELO NETO",
                    "ADRIANA SATURNINO DO NASCIMENTO",
                    "EDMAR CICERO MOREIRA",
                    "JURACI DA COSTA E SILVA",
                    "JUSCELINO SANTOS CRUZ",
                    "RAFAEL VALDIVINO DOS SANTOS",
                    "MARCOS ROBERTO DA SILVA",
                    "ALTAMIRO FERREIRA SICONHA",
                    "LEANDRO DA SILVA SILVEIRA",
                    "SAMUEL PEREIRA DA SILVA",
                    "ALISON HENRIQUE MAIOLLI FEITOZA",
                    "ADRIANO NOGUEIRA",
                    "ROSANGELA BIZERRA DO NASCIMENTO",
                    "JEFFERSON RODRIGO DE FREITAS FRASSON",
                    "LEO PEREIRA DE SOUZA",
                    "WESLER ANDERSON BIAZIM",
                    "RONALDO HENRIQUE CRISTO TEIXEIRA",
                    "CLAUDOMIRO APARECIDO PAULO",
                    "RODRIGO OLIVEIRA GASPARETTI",
                    "TANIA PEREIRA BARBOSA LOPES",
                    "FERNANDO FONTANELE DOS SANTOS",
                    "VIVIANE ANDREZA DOS SANTOS",
                    "FELIPE DE OLIVEIRA SOUZA",
                    "RAFAEL BOTELHO ALBERTINO",
                    "TIAGO BARRETO ALMEIDA",
                    "CLAUDECIR APARECIDO MIRANDA",
                    "ADRIANA CANDIDO DO NASCIMENTO",
                    "MARIO JOSE CONRAD",
                    "RINALDO RIBEIRO ARAUJO",
                    "LUIZ FERNANDO NEVES DA CRUZ",
                    "JOSE CICERO DE OLIVEIRA",
                    "ANTONIO MARCOS RODRIGUES",
                    "RAFAEL CAETANO DA COSTA",
                    "NADSON DA SILVA ANDRADE",
                    "VALTER SOARES DE MIRANDA",
                    "ALEXANDRE DE SOUZA",
                    "MARCIA REGINA DO NASCIMENTO",
                    "LUIZ CARLOS GALVAO",
                    "GILBERTO APARECIDO DA SILVA",
                    "DECQUISON OLIVEIRA SANTOS",
                    "EDMAR PEREIRA DE SOUZA",
                    "MARCOS GONÇALVES LEANDRO",
                    "JOAO LIMA DE ARAUJO",
                    "ADRIANA CRISTINA LIMA DA SILVA",
                    "VANI DA SILVA",
                    "RENATO APARECIDO DA SILVA SOUZA",
                    "AMARILDO XAVIER DE OLIVEIRA",
                    "RUBENS UBERNA",
                    "ADENILSON DE SOUZA MACHADO",
                    "ANDRE ROCHA PASSOS",
                    "ADALBERTO DOS SANTOS CARDOSO",
                    "RONNIE ROGERIO MARQUES",
                    "RONALDO KOVACS",
                    "JOANES LEMOS DA SILVA",
                    "ALEXSANDER TEIXEIRA DA SILVA",
                    "GERSON CALLEGARI",
                    "IVAN ALVES VIEIRA",
                    "CELSO MARTINS DE CARVALHO CAMARA",
                    "JORGE HENRIQUE DE OLIVEIRA",
                    "RAFAEL LOPES PEREIRA",
                    "RONALDO RAMOS CASTANHA",
                    "VLADER DANIEL SANTANA DE CAMPOS",
                    "MAURO GOMES",
                    "NILSON EDNEI JULIAO",
                    "MARCELO DE LIZ CAMARGO",
                    "FERNANDO VITOR NASCIMENTO",
                    "IZAIAS PEREIRA DA SILVA",
                    "EDUARDO AKIO CAVALCANTI",
                    "CLAUDEMIR PATEZ FERREIRA",
                    "CLAUDIO DONIZETE DE MORAES",
                    "VERONICA OLIVEIRA DE LEO",
                    "DANIEL AUGUSTO CALIXTO DOS SANTOS",
                    "JENER GARCIA",
                    "GUEDRYAN DE LIMA BORGES",
                    "FERNANDO LUIZ ROSSATO",
                    "EDVALDO DA MOTA",
                    "LUIZ DO CARMO PEIXOTO",
                    "ALEX SILVA ARAUJO",
                    "WAGNER ROBERTO SALDINI QUALIO",
                    "JUNIOR PEREIRA SILVA",
                    "MAURYLIO MOZART SILVA FIORINI",
                    "NEMAR JOSE NOGUEIRA JUNIOR",
                    "KLEBER DOMENICHELI",
                    "ALAN CRISTIAN DE LIMA SILVA",
                    "DIEGO PEREIRA DOS SANTOS",
                    "ADRIANO DE JESUS MOREIRA",
                    "ALEXANDRE LIMA DE SOUZA",
                    "ANTONIO PEREIRA SALES FILHO",
                    "THIAGO GONÇALVES RODRIGUES",
                    "REINALDO APARECIDO DE OLIVEIRA QUINTAL",
                    "PAULO DA SILVA",
                    "FABIO TEODORO DOS SANTOS",
                    "CAINA GODINHO SETRA",
                    "CARLOS ROBERTO ADAO",
                    "CELSO LUIZ DE LIMA",
                    "MARIA SINELZA MOTA SANTOS",
                    "EDNILSON APARECIDO DE SOUZA",
                    "MARCIO ROGERIO GOMES",
                    "LEOCADIO PEREIRA MENDES",
                    "WELLINGTON RODOLFO MARTINS",
                    "VALDEMAR MARQUES FERREIRA",
                    "FELIPE RODRIGUES",
                    "ELTON DE MOURA TEIXERA",
                    "JOSE ELIAS DA SILVA",
                    "RICARDO METZGER RAMOS",
                    "DOUGLAS PIMENTA DE OLIVEIRA",
                    "VANDERLEY JOSÉ BRAGA",
            };
        }
    }
}