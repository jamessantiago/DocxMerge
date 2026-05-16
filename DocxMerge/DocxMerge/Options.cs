using CommandLine;

namespace DocxMerge;

class Options
{
    [Option('i', "input", Required = true, HelpText = "A list of docx files to merge in order")]
    public IEnumerable<string> InputFiles { get; set; }

    [Option('o', "output", Required = false, HelpText = "Output file [Default: output.docx]")]
    public string Output { get; set; }

    [Option('f', "Force", Required = false, HelpText = "Replace output if already exists")]
    public bool Force { get; set; }

    [Option('r', "Repair-Spacing", Required = false, HelpText = "Replace single spacing after sentences removed from pandoc")]
    public bool RepairSpacing { get; set; }

    [Option('v', "Verbose", Required = false, HelpText = "Show more information when executing")]
    public bool Verbose { get; set; }
}
