using System;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: StripOfficeMetadata <file>");
            return;
        }

        string filePath = args[0];
        string ext = Path.GetExtension(filePath).ToLower();

        switch (ext)
        {
            case ".docx":
                StripWordMetadata(filePath);
                break;
            case ".xlsx":
                StripExcelMetadata(filePath);
                break;
            case ".pptx":
                StripPowerPointMetadata(filePath);
                break;
            default:
                Console.WriteLine("Unsupported file type: " + ext);
                break;
        }
    }

    static void StripWordMetadata(string path)
    {
        using (var doc = WordprocessingDocument.Open(path, true))
        {
            // Remove comments
            doc.MainDocumentPart?.DeleteParts(doc.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>());

            // Remove tracked changes by stripping markup
            foreach (var ins in doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.InsertedRun>())
                ins.Remove();
            foreach (var del in doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.DeletedRun>())
                del.Remove();

            doc.MainDocumentPart.Document.Save();
        }

        RemoveCoreProperties(path);
        Console.WriteLine("Cleaned Word metadata: " + path);
    }

    static void StripExcelMetadata(string path)
    {
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            // Clear workbook properties
            var wbProps = doc.WorkbookPart?.Workbook.WorkbookProperties;
            if (wbProps != null)
            {
                wbProps.CodeName = null;
            }

            doc.WorkbookPart?.Workbook.Save();
        }

        RemoveCoreProperties(path);
        Console.WriteLine("Cleaned Excel metadata: " + path);
    }

    static void StripPowerPointMetadata(string path)
    {
        using (var doc = PresentationDocument.Open(path, true))
        {
            // Remove comment authors if present
            var authorsPart = doc.PresentationPart?.CommentAuthorsPart;
            if (authorsPart != null)
            {
                doc.PresentationPart.DeletePart(authorsPart);
            }

            doc.PresentationPart?.Presentation.Save();
        }

        RemoveCoreProperties(path);
        Console.WriteLine("Cleaned PowerPoint metadata: " + path);
    }

    static void RemoveCoreProperties(string path)
    {
        // Open OPC package directly to wipe metadata
        using (var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite))
        {
            var props = package.PackageProperties;
            props.Creator = null;
            props.LastModifiedBy = null;
            props.Title = null;
            props.Description = null;
            props.Keywords = null;
            props.Subject = null;
            props.Category = null;
            props.ContentStatus = null;
            props.Version = null;
        }
    }
}
