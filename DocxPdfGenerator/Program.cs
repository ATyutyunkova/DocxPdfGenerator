using DocxPdfGenerator.Services;

namespace DocxPdfGenerator;

public static class Program
{
    public static void Main(string[] args)
    {
        Console.WriteLine("Enter a number: 1 - GenerateNewDocx, 2 - GenerateDocxFromXml");
        var doWhat = Console.ReadLine();

        switch (doWhat)
        {
            case "1":
                GenerateNewDocx();
                break;
            case "2":
                GenerateDocxFromXml();
                break;
            default:
                return;
        }
    }

    private static void GenerateNewDocx()
    {
        Console.WriteLine("Enter the file path with name of file and extension '.docx': ");
        var filePath = Console.ReadLine();
        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Error: Empty file path!");
            return;
        }

        Console.WriteLine("Enter the message: ");
        var msg = Console.ReadLine();
        if (string.IsNullOrEmpty(msg))
        {
            msg = "test string";
        }

        OpenXmlGenerator.GenerateNewDocx(filePath, msg);
    }

    private static void GenerateDocxFromXml()
    {
        Console.WriteLine("Enter the XML file path with name of file and extension '.xml': ");
        var xmlFilePath = Console.ReadLine();
        if (string.IsNullOrEmpty(xmlFilePath))
        {
            Console.WriteLine("Error: Empty file path!");
            return;
        }

        Console.WriteLine("Enter the output file path with name of file and extension '.docx': ");
        var outputDocument = Console.ReadLine();

        OpenXmlGenerator.GenerateDocxFromXml(xmlFilePath, outputDocument);
    }
}