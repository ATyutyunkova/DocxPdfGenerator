using DocxPdfGenerator.Services;

namespace DocxPdfGenerator;

public static class Program
{
    public static void Main(string[] args)
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

        OpenXmlGenerator.GenerateDocument(filePath, msg);
    }
}