using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;

class Program
{
    static void Main(string[] args)
    {
        string pastaRaiz = @"C:\Users\murilo.souza\Desktop\Teste - Modo Deus";

        int deletados = 0;
        int erros = 0;

        var opcoes = new ParallelOptions
        {
            MaxDegreeOfParallelism = Environment.ProcessorCount
        };

        try
        {
            var arquivos = Directory.EnumerateFiles(
                pastaRaiz,
                "*",
                SearchOption.AllDirectories
            ).ToList();

            // 1. Identifica pastas que possuem arquivos CONSOLIDADO
            var pastasComConsolidado = arquivos
                .Where(a => Path.GetFileName(a)
                    .IndexOf("CONSOLIDADO", StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(Path.GetDirectoryName)
                .Distinct()
                .ToHashSet();

            // 2. Conta PDFs apenas nessas pastas
            int totalPdfsNasPastas = arquivos
                .Where(a =>
                    pastasComConsolidado.Contains(Path.GetDirectoryName(a)) &&
                    Path.GetExtension(a).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                .Count();

            Console.WriteLine($"📄 Total de PDFs nas pastas com CONSOLIDADO: {totalPdfsNasPastas}");

            // 3. Exclusão paralela
            Parallel.ForEach(arquivos, opcoes, arquivo =>
            {
                try
                {
                    var nome = Path.GetFileName(arquivo);

                    if (nome.IndexOf("CONSOLIDADO", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        File.Delete(arquivo);
                        Interlocked.Increment(ref deletados);
                    }
                }
                catch
                {
                    Interlocked.Increment(ref erros);
                }
            });

            Console.WriteLine($"✅ Deletados: {deletados}");
            Console.WriteLine($"⚠️ Erros: {erros}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro geral: {ex.Message}");
        }
    }
}