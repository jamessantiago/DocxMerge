using CommandLine;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DocxMerge;

static partial class Program
{
    static void Main(string[] args)
    {
        IEnumerable<string> inputFiles = null;
        string output = null;
        bool verbose = false;
        bool force = false;
        bool repairSpacing = false;

        var results = Parser.Default.ParseArguments<Options>(args);
        results.MapResult(options =>
        {
            inputFiles = options.InputFiles;
            output = options.Output ?? "output.docx";
            verbose = options.Verbose;
            force = options.Force;
            repairSpacing = options.RepairSpacing;
            return 0;
        }, _ =>
        {
            Environment.Exit(1);
            return 1; // ._.
        });

        if (inputFiles.Count() < 2) ExitWithError("There must be at least two input files");
        foreach (var file in inputFiles)
        {
            if (!File.Exists(file))
                ExitWithError("Unable to find {0}", file);
        }

        try {
            if (verbose) Console.WriteLine("Creating initial document");
            File.Copy(inputFiles.First(), output, force);

            if (verbose) Console.WriteLine("Opening {0} for writing", output);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(output, true))
            {
                int fileId = 0;
                foreach (var filepath in inputFiles.Skip(1))
                {
                    if (verbose) Console.WriteLine("Adding {0} to {1}", filepath, output);
                    fileId++;
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                    string altChuckId = string.Format("AltChuckId{0}", fileId);
                    var mainPart = doc.MainDocumentPart;
                    var chunk = mainPart.AddAlternativeFormatImportPart(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                        altChuckId);

                    if (repairSpacing)
                        RepairSentenceSpacing(filepath);

                    using (FileStream fileStream = File.Open(filepath, FileMode.Open))
                        chunk.FeedData(fileStream);
                    var altChunk = new XElement(w + "altChunk", new XAttribute(r + "id", altChuckId));
                    var mainDocumentXDoc = GetXDocument(doc);

                    mainDocumentXDoc.Root.Element(w + "body").Elements(w + "p").Last().AddAfterSelf(altChunk);
                    SaveXDocument(doc, mainDocumentXDoc);
                }
            }

            if (verbose) Console.WriteLine("Successfully merged all documents");
        }
        catch (Exception ex)
        {
            if (verbose) ExitWithError(ex.ToString());
            else ExitWithError("DocxMerge failed to process the files: {0}", ex.Message);
        }
    }

    private static void RepairSentenceSpacing(string filepath)
    {
        using WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, true);
        string docText = null;
        using (StreamReader sr = new(wordDoc.MainDocumentPart.GetStream()))
        {
            docText = sr.ReadToEnd();
        }

        Regex regexText = RegexText();
        var refIndex = docText.IndexOf("References");
        if (refIndex > 0)
        {
            var bodyText = docText[..refIndex];
            var references = docText[bodyText.Length..];
            docText = regexText.Replace(bodyText, ".  ${a}") + references;
        }
        else
        {
            docText = regexText.Replace(docText, ".  ${a}");
        }

        using StreamWriter sw = new(wordDoc.MainDocumentPart.GetStream(FileMode.Create));
        sw.Write(docText);
    }

    private static XDocument GetXDocument(WordprocessingDocument myDoc)
    {
        // Load the main document part into an XDocument
        XDocument mainDocumentXDoc;
        using (Stream str = myDoc.MainDocumentPart.GetStream())
        using (XmlReader xr = XmlReader.Create(str))
            mainDocumentXDoc = XDocument.Load(xr);
        return mainDocumentXDoc;
    }

    private static void SaveXDocument(WordprocessingDocument myDoc, XDocument mainDocumentXDoc)
    {
        // Serialize the XDocument back into the part
        using Stream str = myDoc.MainDocumentPart.GetStream(
            FileMode.Create, FileAccess.Write);
        using XmlWriter xw = XmlWriter.Create(str);
        mainDocumentXDoc.Save(xw);
    }

    private static void ExitWithError(string message, params object[] args)
    {
        var defaultColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Red;
        if (args.Length != 0)
            Console.WriteLine(message, args);
        else
            Console.WriteLine(message);
        Console.ForegroundColor = defaultColor;
        Environment.Exit(1);
    }

    [GeneratedRegex(@"\. (?<a>[A-Z])")]
    private static partial Regex RegexText();
}
