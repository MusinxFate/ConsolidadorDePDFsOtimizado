namespace CopiadorDePastas;

internal static class Program
{
    // Preencha os caminhos e os nomes das pastas antes de executar.
    private static readonly string CaminhoOrigem = @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\4. CÓPIAS DOS PROCESSOS - BOOKS";
    private static readonly string CaminhoDestino = @"\\SRV-FS02-BBM\Empresas$\Empresas\VIA VAREJO (GRUPO CASAS BAHIA)\1. CONTROLE\7. REVISÃO DA CP NAS RTS\13. PAFs\19613.7393632022-72 - 6ª COMPENSAÇÃO 07.2022\BOOK";
    private static readonly string[] NomesDasPastas =
    [
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
"MARILENE RODRIGUES DOS SANTOS",
"JOSE ARLINDO DA SILVA",
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
"LUIZ CLAUDIO BERNARDO VALENTIM",
"VANESSA CEZAR HIGASHI DOS SANTOS",
"JOSE EDSON SARAIVA DE BARROS",
"ROGER EYVANS DOS SANTOS MARQUES SOARES",
"RAFAEL PEREIRA DA CRUZ",
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
"MARCELO RODRIGUES DE OLIVEIRA",
"SILVIO RANGEL DA CRUZ",
"JOSE CLAUDIO DE ARAUJO MARIANO SOUZA",
"GLEICYANE APARECIDA DE SOUZA GUIMARAES",
"MICAEL MARQUES BARROS",
"FLAVIO BERNARDINO DE SENA",
"AURELIO NONATO",
"ANDRE DE CASTRO DA SILVA",
"JOICE APARECIDA ARAUJO BARRETO",
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
"JORGE MANOEL DE SOUZA",
"FERNANDA VALESKA AGUIAR",
"MAYLOM HENRIQUE BUK MORAES CRUZ",
"FABIO RODRIGUES BARBOSA",
"SAMUEL GOUVEA MAIA",
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
"LILIAN CRISTINA FILGUEIRA DE LIMA",
"JEFERSON ALMEIDA CAVALHEIRO",
"RAFAEL GONCALVES DA SILVA",
"MARIO SIQUEIRA",
"VANIA LUCIA DE PAIVA",
"JOILSON LUCAS ROSARIO BATISTA",
"CLAUDIO LIMA DE OLIVEIRA",
"VANESSA BARBOSA DE SOUSA SILVA",
"JOSSUE SALES DA SILVA",
"VANDIRA DOS SANTOS ESQUIVEL CONDURU",
"MARCOS MOREIRA BARBOSA",
"CLAUDIA MARCIA NONATO",
"JOAO LUCAS CARVALHO DE BRITO",
"CARLOS ALBERTO ALVES SANTOS",
"VANDIRA DOS SANTOS ESQUIVEL CONDURU",
"GUILHERME HENRIQUE ALVES DA SILVA",
"JOAO VICENTE DA SILVA JUNIOR",
"ELIEL MENDES DA SILVA ASSUNCAO",
"LEILA PAULA BARROS",
"DANAEL FERREIRA",
"ADALBERTO CARLOS",
"CLAUDIO BATISTA DA SILVA",
"FABIO SIMOR GOMES",
"CRISTIANO RODRIGUES DE SOUZA",
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
"FLAVIO DOS SANTOS COTA",
"MONICA AMALIA DOS SANTOS",
"JOSE FRANCISCO SOUSA DA SILVA",
"DARIO PEIXOTO LIMA",
"JOSE MARIO SOARES DA COSTA",
"GESSI JACI CABRAL",
"ANTONIO CARLOS CORREA GONCALVES JUNIOR",
"RICARDO RIBEIRO PESSOA",
"NERONILSON HONORATO AMORIM",
"LUCIANO NEGRAO",
"GUILHERME CELESTINO DINIZ",
"VALDINEI APARECIDO DE LIMA",
"JOSE INACIO DO CARMO JUNIOR",
"WELLINGTON DE OLIVEIRA SILVERIO",
"CRISTIANO COSTA CORREA",
"THIAGO LUIZ DA SILVA BARROSO",
"FABIO SMIN",
"CLEUDIA SOARES DE SALES",
"ROBSON JOSE SPINOLA MORAIS",
"RITA DE CASSIA DA SILVA CRUZ",
"LUCAS DE ALMEIDA PEREIRA",
"RAFAEL DE JESUS VIEIRA",
"WAGNER HERCILIO COMIN",
"CLAUDIO PINHO BRANDAO",
"GABRIELLE SAMPAIO DOS SANTOS",
"ALEXANDRE LIMA DE SOUSA JUNIOR",
"RENATA DE OLIVEIRA",
"BRUNO BARBOSA FARNEZI",
"ROSEMERE RUFINO DIAS",
"DIEGO BITTENCOURT DA COSTA",
"ROGERIO FERREIRA GONCALVES",
"FRANCISCO CLEMERSON CAVALCANTE SILVA",
"FERNANDO SOARES DE FIGUEIREDO",
"DIEGO FABIANO DA CRUZ ALVES",
"KLEBSOM ROGERIO DOS SANTOS ARAUJO",
"LAURA BARBOSA",
"ORTIAGO BENITES DE OLIVEIRA",
"JOSE CARLOS DE ALMEIDA",
"ROBSON DOS SANTOS MUNIZ",
"FELIPE GOMES DA COSTA",
"RODRIGO AGENOR SILVA BARBOSA",
"MARIA DA CRUZ DA COSTA SILVA",
"CARLOS OTAVIO DE ARAUJO",
"NATANAEL FAGUNDES DE LIMA",
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
"LUIZ CARLOS DO ROCIO MACHADO",
"CARLOS HENRIQUE CARPES",
"PATRICIA CAMARGOS SILVA SANTOS",
"WEVERLINE NOBRE BAPTISTA",
"MAURO CESAR ISOLA",
"MAYCON BERGAMO NASCIMENTO OVERNEZ",
"CLEBER RAMOS RODRIGUES",
"SIMONE DE SOUZA ASSUMPCAO",
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
"MOACY LIRA DA SILVA",
"JOAO APARECIDO PUGLIANO",
"JOSIAS DA COSTA BARBOSA",
"LEANDRO LACY RODRIGUES DA SILVA",
"ADRIANO PINTO BARCELOS",
"DEJAVAN PEREIRA DOS SANTOS",
"CRISTIANE DE ARAUJO SIMAO",
"EDIVALDO FARIAS DOS SANTOS",
"ALEXANDER DA SILVA LOPES FIGUEIREDO",
"JOSE FRANCISCO BORGES DOS SANTOS",
"FERNANDO GUIZILINI DA SILVA",
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
"FRANCISCO SANTANA BRITO DA SILVA",
"MANOEL JOSE DOS SANTOS",
"BRUNO RODRIGUES DA SILVA",
"WEBERT ELIZIO MACHADO",
"DIRCEU PEIXOTO PENIDO",
"LILIANE IZABEL DA SILVA",
"ANTONIO CARLOS DE ASSIS GOMES",
"ROBSON CORREA DOS SANTOS",
"RAFAEL ARAUJO PINHEIRO",
"THIAGO ANTONIO OLIVEIRA SILVA",
"MARIO FERNANDO PENA",
"ALEX MENDES FIDELES",
"CLEITON DE FARIA COSTA",
"JULIANA ARAUJO ALVES RODRIGUES",
"ALAN FELIX DE SOUZA",
"FABIO HENRIQUE ALVES",
"ANDRE DE CASTRO DA SILVA",
"ANDERSON GONCALVES",
"JOSE CLAUDINO",
"ALILIANA SILVA DE SOUSA",
"RICARDO ALEXANDRE RODES NEVES",
"EDMAR VIEIRA SOARES",
"RAMON ALVES DE SOUZA",
"ANDERSON ALEXANDRE DE ALMEIDA MAZETTO",
"CARLOS ROSEMBERG LACERDA ALEGRE DOS SANTOS",
"SILAS BALBINO",
"SUELI DANTAS SOUZA",
"FRANCISCO DAS CHAGAS LOPES DA SILVA",
"JAIME JOSE HENRIQUES JUNIOR",
"MARCOS VINICIUS DE SOUZA ALVES",
"EVANDRO BARBOSA LEITE",
"ROBSON ALVES DA SILVA",
"LUIZ FERNANDO SOARES RIBEIRO",
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
"LEONARDO SOARES RIBEIRO",
"FERNANDA VIEIRA DELGADO",
"VALDEMILSON ROSA DOS SANTOS",
"RENATO PACHECO SANTOS",
"DENILSON VICENTE DE OLIVEIRA",
"MARCOS ANTONIO DA SILVA",
"WILK ASSIS OLIVEIRA",
"JOSE RAIMUNDO FILHO",
"RENATA BENONI TEODORO",
"WILK ASSIS OLIVEIRA",
"VANESSA ACBAS MARTINELLI",
"CATIANE OLIVEIRA ANDRADE",
"CRISTIANO DA SILVA LUIZ",
"ROGERIO SILVA SANTOS",
"JOAO BATISTA LEAO NETO",
"CLAUDIO DA SILVA BATISTA",
"GISELLE SILVESTRE DA SILVA",
"SILVIO DE ASSIS MAMEDIO",
"BRUNO FERNANDES OLIVEIRA",
"CLAUDIO CAVALINI GUILHERME",
"MARCELO SANTANA PRADO",
"JOAO BATISTA DE OLIVEIRA",
"VANDO ALVES DOS SANTOS",
"LUCIANA CAETANO BRAGA",
"NEY MARQUES MARTINS",
"MISAEL PEDRO DA SILVA MARCELINO",
"ALESSANDRA APARECIDA DA SILVA",
"CLAUDIO MARCIO DE BRITO",
"SANDRA GUERRA DOS SANTOS DEVICSI",
"CHRISTIANE GONCALVES FARIAS",
"ISAQUE MACHADO DE SOUSA",
"THIAGO ALVES DA SILVA GOMES",
"JOAO RAFAEL SOARES DOS REIS",
"AMILTON PEREIRA DE AZEREDO JUNIOR",
"ALFRANDE FERRAZ DE LIMA FILHO",
"SIDNEI ALVES DE MELLO",
"MONICA FERNANDES ANDRADE",
"THIAGO DA SILVA COELHO",
"IZAIAS RODRIGUES BUENO",
"MAURICIO MARTINS SOUZA",
"JOSE ARAUJO DE ALMEIDA",
"LAISA CORREA DA SILVA",
"JOSE VILSON CARDOSO DA SILVA",
"CLAUDIA SOLIDADE SOUZA SILVA",
"SUZANA AUGUSTA DE LOURENCO",
"ZENAIDE DE ANDRADE CHAVES",
"BRUNA LUIZA DE ALMEIDA OLIVEIRA",
"ROBSON LIMA DE SOUZA",
"CARLOS ALBERTO LEONI",
"ARIOVALDO ROBERTO DA SILVA",
"MARCIO VAZ DE LIMA",
"JOSE ROSA DA SILVA",
"EZEQUIAS DOMINGUES RIBEIRO",
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
"JANICE TERESINHA ALVES DE CARVALHO KRAUS",
"EMILENE RODRIGUES",
"FELIPE DA COSTA REIS",
"KARINA TAMIKO TIAGO TSUBOI",
"DOUGLAS FERREIRA CARNEIRO",
"JUNIOR SERGIO HERMOGENES ARAUJO",
"EDIL GUILHERME DE ALMEIDA",
"DANIELE RIBEIRO NARDON",
"ALMINA DIAS LUZ",
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
"MARCELO PIRIS VARELA",
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
"MARCOS PEREIRA DOS SANTOS",
"JOSE ELISSON BARBOSA DE LIMA",
"ANGELITA DA SILVA MAGALHAES",
"THIAGO PALOMO FIRMA",
"ELISANGELA MARIA ZAGARE RODRIGUES",
"ELISANGELA MARIA ZAGARE RODRIGUES",
"MAYCON RODRIGUES DA ROCHA",
"BENEDITO PEREIRA SOARES",
"DALCIO GOULART VIEIRA",
"JULIANA FERREIRA DE PAULA",
"CARLOS EDUARDO DE OLIVEIRA",
"FRANCISCO SANTANA BRITO DA SILVA i",
"MARCIO JOSE DOS SANTOS",
"CHRISTIANE GONCALVES FARIAS",
"GLEISON ROBERT SANTANA",
"MICHAEL MENDONCA DE SOUZA",
"ADEMIR TOME",
"FABRICIO CAMPOS NOGUEIRA SANTOS",
"ALEXANDER VIANA BRANDAO",
"ADEMIR TOME",
"ADEMIR TOME",
"ROSEMERE DE SOUZA SILVA",
"JOSE RUI ZAGO",
"JULIO MARIA DE MELO",
"GEDSON JUNIOR DE SOUSA SILVA",
"PAULO SERGIO DA SILVA",
"CARLOS ANDRE DE JESUS",
"JOSE JUNHO PEREIRA GOUVEIA",
"JACI RAIMUNDO DANTAS",
"LEANDRO DA SILVA RODRIGUES",
"ARNALDO GUERRA DE SOUZA",
"JONATAS MARTINS TAVARES",
"LUDIMYLLA ALVES DE SOUSA",
"MAICON SANDRO VEIGA RANGEL"
    ];

    private const int MaximoDeCopiasSimultaneas = 16;
    private const bool SobrescreverArquivosExistentes = true;
    // Por padrão, salva os logs na pasta do executável. Pode ser alterado.
    private static readonly string CaminhoDosLogs = Path.Combine(AppContext.BaseDirectory, "logs");

    private static int Main()
    {
        try
        {
            using var log = new LogExecucao(CaminhoDosLogs);
            var cronometro = System.Diagnostics.Stopwatch.StartNew();
            var resultado = 1;
            log.Escrever("INFO", $"Início da execução. Log: {log.Caminho}");
            log.Escrever("INFO", $"Origem configurada: '{CaminhoOrigem}'. Destino configurado: '{CaminhoDestino}'.");
            log.Escrever("INFO", $"Paralelismo: {MaximoDeCopiasSimultaneas}. Sobrescrever existentes: {SobrescreverArquivosExistentes}.");
            log.Escrever("INFO", "Links simbólicos e junções são ignorados durante a enumeração.");

            try
            {
                resultado = CopiarPastas(
                    CaminhoOrigem, CaminhoDestino, NomesDasPastas,
                    MaximoDeCopiasSimultaneas, SobrescreverArquivosExistentes, log);
                return resultado;
            }
            catch (Exception ex)
            {
                log.Escrever("ERRO", "Falha que impediu a conclusão da cópia.", ex);
                return 1;
            }
            finally
            {
                log.Escrever("INFO", $"Fim da execução. Estado: {(resultado == 0 ? "SUCESSO" : "COM FALHAS")}. " +
                    $"Duração: {cronometro.Elapsed:c}. Código de saída: {resultado}.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Não foi possível criar, gravar ou finalizar o log em '{CaminhoDosLogs}': {ex}");
            return 1;
        }
    }

    private static int CopiarPastas(
        string origem, string destino, string[] nomes, int paralelismo, bool sobrescrever, LogExecucao log)
    {
        if (string.IsNullOrWhiteSpace(origem) || string.IsNullOrWhiteSpace(destino))
            throw new ArgumentException("Preencha CaminhoOrigem e CaminhoDestino no início do arquivo.");

        if (paralelismo < 1)
            throw new ArgumentException("MaximoDeCopiasSimultaneas deve ser maior que zero.");

        var nomesSolicitados = nomes
            .Where(nome => !string.IsNullOrWhiteSpace(nome))
            .Select(nome => nome.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (nomesSolicitados.Count == 0)
            throw new ArgumentException("Preencha o array NomesDasPastas no início do arquivo.");

        log.Escrever("INFO", $"Nomes válidos e únicos solicitados: {nomesSolicitados.Count}. " +
            $"Entradas vazias ou duplicadas desconsideradas: {nomes.Length - nomesSolicitados.Count}.");

        origem = Path.TrimEndingDirectorySeparator(Path.GetFullPath(origem));
        destino = Path.TrimEndingDirectorySeparator(Path.GetFullPath(destino));
        log.Escrever("INFO", $"Caminhos completos: origem '{origem}'; destino '{destino}'.");

        // Evita copiar para a própria origem ou misturar árvores sobrepostas.
        if (EstaDentroOuIgual(destino, origem) || EstaDentroOuIgual(origem, destino))
            throw new ArgumentException("Origem e destino devem ser pastas distintas, sem uma dentro da outra.");

        if (!Directory.Exists(origem))
            throw new DirectoryNotFoundException($"Origem não encontrada ou inacessível: {origem}");

        var opcoesEnumeracao = new EnumerationOptions
        {
            AttributesToSkip = FileAttributes.ReparsePoint,
            IgnoreInaccessible = false
        };

        // Apenas as pastas diretamente dentro da origem são comparadas com o array.
        var pastas = new DirectoryInfo(origem)
            .EnumerateDirectories("*", opcoesEnumeracao)
            .Where(pasta => nomesSolicitados.Contains(pasta.Name.Trim()))
            .ToArray();

        var nomesEncontrados = pastas.Select(pasta => pasta.Name.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var naoEncontrados = nomesSolicitados.Except(nomesEncontrados, StringComparer.OrdinalIgnoreCase).ToArray();

        log.Escrever("INFO", $"Pastas encontradas para cópia: {pastas.Length}. Pastas não encontradas: {naoEncontrados.Length}.");

        foreach (var nome in naoEncontrados)
            log.Escrever("AVISO", $"Pasta não encontrada na origem '{origem}': '{nome}'.");

        if (pastas.Length == 0)
        {
            log.Escrever("ERRO", "Nenhuma pasta correspondente foi encontrada na origem. Nenhum arquivo foi copiado.");
            return 1;
        }

        Directory.CreateDirectory(destino);
        var arquivosCopiados = 0;
        var arquivosIgnorados = 0;
        var erros = 0;
        var arquivosComFalha = 0;
        var diretoriosComFalha = 0;
        var pastasConcluidas = 0;

        log.Escrever("INFO", $"Copiando {pastas.Length} pasta(s), com até {paralelismo} cópias simultâneas.");

        Parallel.ForEach(pastas, new ParallelOptions { MaxDegreeOfParallelism = paralelismo }, pasta =>
        {
            var houveErro = false;
            var tempoPasta = System.Diagnostics.Stopwatch.StartNew();
            log.Escrever("INFO", $"Iniciando pasta: '{pasta.FullName}' -> '{Path.Combine(destino, pasta.Name)}'.");
            var pendentes = new Stack<(string Origem, string Destino)>();
            // Preserva o nome original da pasta no destino.
            pendentes.Push((pasta.FullName, Path.Combine(destino, pasta.Name)));

            while (pendentes.TryPop(out var atual))
            {
                try
                {
                    Directory.CreateDirectory(atual.Destino);

                    foreach (var subpasta in Directory.EnumerateDirectories(atual.Origem, "*", opcoesEnumeracao))
                        pendentes.Push((subpasta, Path.Combine(atual.Destino, Path.GetFileName(subpasta))));

                    foreach (var arquivo in Directory.EnumerateFiles(atual.Origem, "*", opcoesEnumeracao))
                    {
                        var arquivoDestino = Path.Combine(atual.Destino, Path.GetFileName(arquivo));
                        try
                        {
                            if (!sobrescrever && File.Exists(arquivoDestino))
                            {
                                Interlocked.Increment(ref arquivosIgnorados);
                                log.Escrever("INFO", $"Arquivo existente preservado: '{arquivo}' -> '{arquivoDestino}'.", exibirNoConsole: false);
                                continue;
                            }

                            File.Copy(arquivo, arquivoDestino, overwrite: sobrescrever);
                            Interlocked.Increment(ref arquivosCopiados);
                            log.Escrever("INFO", $"Arquivo copiado: '{arquivo}' -> '{arquivoDestino}'.", exibirNoConsole: false);
                        }
                        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                        {
                            houveErro = true;
                            Interlocked.Increment(ref erros);
                            Interlocked.Increment(ref arquivosComFalha);
                            log.Escrever("ERRO", $"Não foi possível copiar o arquivo: '{arquivo}' -> '{arquivoDestino}'.", ex);
                        }
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    houveErro = true;
                    Interlocked.Increment(ref erros);
                    Interlocked.Increment(ref diretoriosComFalha);
                    log.Escrever("ERRO", $"Não foi possível processar o diretório: '{atual.Origem}' -> '{atual.Destino}'. " +
                        "Parte do conteúdo pode não ter sido enumerada ou copiada.", ex);
                }
            }

            if (houveErro)
                log.Escrever("ERRO", $"Pasta finalizada com erros: '{pasta.Name}'. Duração: {tempoPasta.Elapsed:c}.");
            else
            {
                Interlocked.Increment(ref pastasConcluidas);
                log.Escrever("INFO", $"Pasta concluída: '{pasta.Name}'. Duração: {tempoPasta.Elapsed:c}.");
            }
        });

        log.Escrever("INFO", $"RESUMO | Pastas concluídas sem erros: {pastasConcluidas}/{pastas.Length}; " +
            $"pastas com erros: {pastas.Length - pastasConcluidas}; pastas não encontradas: {naoEncontrados.Length}; " +
            $"arquivos copiados: {arquivosCopiados}; arquivos existentes preservados: {arquivosIgnorados}; " +
            $"arquivos com falha: {arquivosComFalha}; diretórios com falha: {diretoriosComFalha}; total de erros: {erros}.");
        return erros == 0 && naoEncontrados.Length == 0 ? 0 : 1;
    }

    private static bool EstaDentroOuIgual(string caminho, string pasta)
    {
        var prefixo = Path.EndsInDirectorySeparator(pasta) ? pasta : pasta + Path.DirectorySeparatorChar;
        return caminho.Equals(pasta, StringComparison.OrdinalIgnoreCase)
               || caminho.StartsWith(prefixo, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class LogExecucao : IDisposable
    {
        private readonly object sincronizacao = new();
        private readonly StreamWriter arquivo;

        public string Caminho { get; }

        public LogExecucao(string diretorio)
        {
            Directory.CreateDirectory(diretorio);
            Caminho = Path.GetFullPath(Path.Combine(diretorio,
                $"copia_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}.log"));
            arquivo = new StreamWriter(new FileStream(Caminho, FileMode.CreateNew, FileAccess.Write, FileShare.Read),
                new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: true))
            {
                AutoFlush = true
            };
        }

        public void Escrever(string nivel, string mensagem, Exception? excecao = null, bool exibirNoConsole = true)
        {
            // Mantém cada registro completo, mesmo com várias cópias simultâneas.
            lock (sincronizacao)
            {
                var registro = $"[{DateTimeOffset.Now:O}] [{nivel}] {mensagem}";
                if (excecao is not null)
                    registro += Environment.NewLine + excecao;

                arquivo.WriteLine(registro);
                if (exibirNoConsole)
                {
                    if (nivel is "ERRO" or "AVISO")
                        Console.Error.WriteLine(registro);
                    else
                        Console.WriteLine(registro);
                }
            }
        }

        public void Dispose()
        {
            lock (sincronizacao)
                arquivo.Dispose();
        }
    }
}
