using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxPdfGenerator.Services;

public static class OpenXmlGenerator
{
    public static void GenerateDocument(string filepath, string msg)
    {
        using (var doc =
               WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add a main document part
            var mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());

            // String msg contains the text
            run.AppendChild(new Text(msg));
        }
    }
}