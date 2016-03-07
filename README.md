# DocxMerge
Simple command line program to merge word documents

## Getting Started

First [grab a copy](https://github.com/jamessantiago/DocxMerge/releases) of DocxMerge.exe and copy to your local system (preferably in your [PATH](https://en.wikipedia.org/wiki/PATH_%28variable%29) for easy access).

Run this command:

    DocxMerge -i MyFirstDocument.docx MySecondDocument.docx

This will merge your two documents in the order that you listed the files and output the results to output.docx.

The full option listing can be seen using the `--help` switch:

    > .\DocxMerge.exe --help
    DocxMerge 1.1.0.0
    Copyright c James Santiago 2015

      -i, --input             Required. A list of docx files to merge in order

      -o, --output            Output file [Default: output.docx]

      -f, --Force             Replace output if already exists

      -r, --Repair-Spacing    Replace single spacing after sentences removed from
                              pandoc

      -v, --Verbose           Show more information when executing

      --help                  Display this help screen.

      --version               Display version information.
