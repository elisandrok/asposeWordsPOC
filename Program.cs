using Aspose.Words;
using Aspose.Words.Replacing;

namespace asposeWordsPOC
{
    class Program
    {
        static void Main(string[] args)
        {
            License license = new License();
            try
            {
                license.SetLicense("Aspose.Words.lic");
                
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }

            Document docA = new Document();            
            DocumentBuilder builder = new DocumentBuilder(docA);

            builder.MoveToDocumentStart();
            //builder.Write("Primeiro teste com relatórios no Aspose Words");

            Document docB = new Document("C:\\temp\\TemplatesAspose\\Document.docx");
            docA.AppendDocument(docB, ImportFormatMode.KeepSourceFormatting);

            // Condições para determinar se as seções devem ser removidas ou não
            bool imprimirSecao1 = true;  // Defina a condição aqui
            bool imprimirSecao2 = true;  // Defina a condição aqui

            // Manipular as seções baseadas nas condições
            if (!imprimirSecao1)
            {
                Bookmark secao1 = docA.Range.Bookmarks["Secao1"];
                secao1.Text = string.Empty; // Remove o conteúdo da seção
            }
            if (!imprimirSecao2)
            {
                Bookmark secao2 = docA.Range.Bookmarks["Secao2"];
                secao2.Text = string.Empty; // Remove o conteúdo da seção
            }            

            Dictionary<string, string> placeholders = new Dictionary<string, string>
            {
                { "<<NomeCliente>>", "João Silva" },
                { "<<ResEndereco>>", "Rua Teste" },
                { "<<Idade>>", "30" },
                { "<<Data>>", DateTime.Now.ToString("dd/MM/yyyy") }
            };

            // Substituir os placeholders
            foreach (var placeholder in placeholders)
            {
                docA.Range.Replace(placeholder.Key, placeholder.Value, new FindReplaceOptions());
            }

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Página ");
            builder.InsertField("PAGE", "");
            builder.Write(" de ");
            builder.InsertField("NUMPAGES", "");

            Document subRelatorio = new Document("C:\\temp\\TemplatesAspose\\SubReport.docx");
            Dictionary<string, string> placeholdersSub = new Dictionary<string, string>
            {
                { "<<NomeCliente>>", "João Silva" },
                { "<<ResEndereco>>", "Rua Teste" },
                { "<<Idade>>", "30" },
                { "<<Data>>", DateTime.Now.ToString("dd/MM/yyyy") }
            };

            // Substituir os placeholders
            foreach (var placeholder in placeholdersSub)
            {
                subRelatorio.Range.Replace(placeholder.Key, placeholder.Value, new FindReplaceOptions());
            }
            docA.AppendDocument(subRelatorio, ImportFormatMode.KeepSourceFormatting);
            
            docA.Save("C:\\temp\\TemplatesAspose\\Report.pdf");
        }
    }
}

