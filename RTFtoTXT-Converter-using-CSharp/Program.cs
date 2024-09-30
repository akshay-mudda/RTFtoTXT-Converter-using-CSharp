using System;
using System.IO;
using System.Configuration;
using RtfPipe;
using HtmlAgilityPack;

namespace RTFtoTXT_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Read source and destination paths from app.config
            string sourcePath = ConfigurationManager.AppSettings["SourcePath"];
            string destinationPath = ConfigurationManager.AppSettings["DestinationPath"];

            // Check if destination folder exists, create if not
            if (!Directory.Exists(destinationPath))
            {
                Directory.CreateDirectory(destinationPath);
            }

            // Get all RTF files in the source folder
            string[] rtfFiles = Directory.GetFiles(sourcePath, "*.rtf");

            if (rtfFiles.Length == 0)
            {
                Console.WriteLine("No RTF files found in the source folder.");
                return;
            }

            foreach (var rtfFilePath in rtfFiles)
            {
                try
                {
                    // Get file name without extension
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(rtfFilePath);
                    string txtFilePath = Path.Combine(destinationPath, fileNameWithoutExtension + ".txt");

                    // Check if the TXT already exists
                    if (File.Exists(txtFilePath))
                    {
                        Console.WriteLine($"TXT already exists for {fileNameWithoutExtension}. Skipping file.");
                        continue;
                    }

                    // Convert RTF file to TXT
                    ConvertRtfToTxt(rtfFilePath, txtFilePath);

                    // After conversion, delete the source RTF file
                    File.Delete(rtfFilePath);

                    Console.WriteLine($"Successfully converted {fileNameWithoutExtension}.rtf to TXT.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {rtfFilePath}: {ex.Message}");
                }
            }
        }

        // Convert RTF file to TXT (RTF -> HTML -> TXT)
        static void ConvertRtfToTxt(string rtfFilePath, string txtFilePath)
        {
            // Read the RTF content from the file
            string rtfContent = File.ReadAllText(rtfFilePath);

            // Convert RTF to HTML using RtfPipe
            string htmlContent = Rtf.ToHtml(rtfContent);

            // Extract plain text from the HTML content using HtmlAgilityPack
            string plainText = ConvertHtmlToText(htmlContent);

            // Write the plain text to the destination TXT file
            File.WriteAllText(txtFilePath, plainText);
        }

        // Extract plain text from HTML using HtmlAgilityPack
        static string ConvertHtmlToText(string htmlContent)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);

            // Extract the inner text, which is the plain text representation of the HTML content
            return doc.DocumentNode.InnerText;
        }
    }
}