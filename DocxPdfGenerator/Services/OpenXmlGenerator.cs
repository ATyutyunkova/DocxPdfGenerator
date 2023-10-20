using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxPdfGenerator.Services;

public static class OpenXmlGenerator
{
    public static void GenerateNewDocx(string filepath, string msg)
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

    public static void GenerateDocxFromXml(string xmlFilePath, string outputDocument)
    {
        var templateDocument =
            @"C:\Users\tyutyunkova\source\altecReposRider\DocxPdfGenerator\DocxPdfGenerator\Resources\test\ExampleDoc.docx";

        using (var wordDoc = WordprocessingDocument.Open(templateDocument, true))
        {
            //get the main part of the document which contains CustomXMLParts
            var mainPart = wordDoc.MainDocumentPart;

            //delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable
            mainPart.DeleteParts<CustomXmlPart>(mainPart.CustomXmlParts);

            //add new CustomXMLPart with data from new XML file
            var myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var stream = new FileStream(xmlFilePath, FileMode.Open))
            {
                myXmlPart.FeedData(stream);
            }

            wordDoc.Clone(outputDocument);
        }
    }
}