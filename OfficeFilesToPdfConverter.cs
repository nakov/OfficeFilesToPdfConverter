
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;

using Application = Microsoft.Office.Interop.Word.Application;
using _Application = Microsoft.Office.Interop.Word._Application;

public class Program
{
    private const string DocxExtention = ".docx";
    private const string DocExtention = ".doc";
    private const string PptExtention = ".ppt";
    private const string PptxExtention = ".pptx";
    private const string PdfExtention = ".pdf";

    internal static void Main()
    {
        Console.WriteLine("Office to PDF converter starting...");

        var allFiles = new List<FileInfo>();
        DirectoryInfo dirInfo = new DirectoryInfo(@".\");

        EnumerateFilesRecursively(dirInfo, allFiles);

        var word2010Files = allFiles.Where(x => x.FullName.EndsWith(DocxExtention)).ToList();
        var word2007Files = allFiles.Where(x => x.FullName.EndsWith(DocExtention)).ToList();
        var pptFiles = allFiles.Where(x => x.FullName.EndsWith(PptExtention)).ToList();
        var pptxFiles = allFiles.Where(x => x.FullName.EndsWith(PptxExtention)).ToList();

        Console.WriteLine(
            "Office files found for converting: docx = {0}, doc = {1}, ppt = {2}, pptx = {3}",
            word2010Files.Count,
            word2007Files.Count,
            pptFiles.Count,
            pptxFiles.Count);

        ConvertWordToPdf(word2010Files, DocxExtention, PdfExtention);
        ConvertWordToPdf(word2007Files, DocExtention, PdfExtention);
        ConvertPresentationToPdf(pptFiles, PptExtention, PdfExtention);
        ConvertPresentationToPdf(pptxFiles, PptxExtention, PdfExtention);
    }

    private static void EnumerateFilesRecursively(DirectoryInfo dirInfo, List<FileInfo> files)
    {
        files.AddRange(dirInfo.GetFiles());
        var directories = dirInfo.GetDirectories();
        foreach (var directoryInfo in directories)
        {
            EnumerateFilesRecursively(directoryInfo, files);
        }
    }

    private static void ConvertWordToPdf(
        IList<FileInfo> files,
        string inputFileExtention,
        string outputFileExtention)
    {
        if (files.Count == 0)
            return;

        // Create a new Microsoft Word application object
        var word = new Application
        {
            Visible = false,
            ScreenUpdating = false
        };

        foreach (FileInfo file in files)
        {
            try
            {
                // Cast as Object for word Open method
                object filename = (object)file.FullName;

                PrintWarrningMessage(string.Format("Converting to PDF: {0}", filename));

                // Use the dummy value as a placeholder for optional arguments
                Document doc = word.Documents.Open(ref filename);
                doc.Activate();

                object outputFileName = file.FullName.Replace(
                    inputFileExtention, outputFileExtention);
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(
                    ref outputFileName,
                    ref fileFormat);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges);
                doc = null;

                PrintSuccessMessage(string.Format("Done: {0}", outputFileName));
            }
            catch (Exception)
            {
                PrintErrorMessage(string.Format("Convertion failed!"));
            }
        }

        // Cast MS Word to type _Application, so that it will find the correct Quit method
        ((_Application)word).Quit();
        word = null;
        GC.Collect();
    }

    private static void ConvertPresentationToPdf(
        IList<FileInfo> files,
        string inputFileExtention,
        string outputFileExtention)
    {
        if (files.Count == 0)
            return;

        Microsoft.Office.Interop.PowerPoint.Application powerPointApp = 
            new Microsoft.Office.Interop.PowerPoint.Application();

        foreach (var file in files)
        {
            try
            {
                PrintWarrningMessage(string.Format("Converting to PDF: {0}", file.FullName));

                Presentation presentation = powerPointApp.Presentations.Open(
                file.FullName,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoFalse);

                var outputFilename = file.FullName.Replace(inputFileExtention, outputFileExtention);

                presentation.ExportAsFixedFormat(
                    outputFilename,
                    PpFixedFormatType.ppFixedFormatTypePDF,
                    PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                    PpPrintOutputType.ppPrintOutputSlides,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    null,
                    PpPrintRangeType.ppPrintAll,
                    string.Empty,
                    false,
                    false,
                    false,
                    true,
                    true);

                presentation.Close();
                presentation = null;

                PrintSuccessMessage(string.Format("Done: {0}", outputFilename));
            }
            catch (Exception)
            {
                PrintErrorMessage(string.Format("Convertion failed!"));
            }
        }

        powerPointApp.Quit();
        powerPointApp = null;
        GC.Collect();
    }

    private static void PrintWarrningMessage(string content)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(content);
        Console.ResetColor();
    }

    private static void PrintSuccessMessage(string content)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(content);
        Console.ResetColor();
    }

    private static void PrintErrorMessage(string content)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(content);
        Console.ResetColor();
    }
}
