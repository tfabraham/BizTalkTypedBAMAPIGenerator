// BizTalk Typed BAM API Generator
// Copyright (C) 2008-Present Thomas F. Abraham. All Rights Reserved.
// Licensed under the MIT License. See License.txt in the project root.

using System;
using System.IO;
using System.Text;

namespace ExportBamDefinitionXml
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return -1;
            }

            string xLSFileName = args[0];
            string outputFileName = args[1];
            bool useAutomation = false;

            if (args.Length == 3)
            {
                // If parses to true then useAutomation will be true, otherwise it will be false.
                bool.TryParse(args[2], out useAutomation);
            }

            PrintHeader();

            try
            {
                if (useAutomation)
                {
                    Console.WriteLine("Exporting in legacy Excel Automation mode.");
                    Console.WriteLine();
                }

                Console.Write("Exporting BAM XML Definition from the Excel Spreadsheet... ");
                string bAMDefinitionXML = Shared.BamDefinitionXmlExporter.GetBamDefinitionXml(xLSFileName, useAutomation);
                File.WriteAllText(outputFileName, bAMDefinitionXML, UnicodeEncoding.Unicode);
                Console.WriteLine("Success");
                Console.WriteLine();
                Console.WriteLine("Wrote BAM XML to '{0}'.", outputFileName);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Failed");
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(exception.Message);
                Console.ResetColor();
                return -1;
            }

            return 0;
        }

        private static void PrintHeader()
        {
            Console.ForegroundColor = ConsoleColor.White;

            Version assemblyVersion = System.Reflection.Assembly.GetEntryAssembly().GetName().Version;
            Console.WriteLine(
                "BizTalk BAM Definition XML Exporter "
                + assemblyVersion.Major + "." + assemblyVersion.Minor + "." + assemblyVersion.Build);
            Console.WriteLine("[https://github.com/tfabraham/BizTalkTypedBAMAPIGenerator/]");
            Console.WriteLine("Copyright (C) 2007-08 Darren Jefford and 2008-20 Thomas F. Abraham");
            Console.WriteLine();

            Console.ResetColor();
        }

        private static void PrintUsage()
        {
            PrintHeader();
            Console.WriteLine("Exports the XML from an Excel BAM Definition workbook");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("ExportBamDefinitionXml <ExcelFile> <XMLFile> [UseLegacyExport]");
            Console.WriteLine(@"e.g: ExportBamDefinitionXml C:\BAMDefinition.xls C:\BAMDefinition.xml");
            Console.WriteLine();
            Console.WriteLine("[UseLegacyExport] (Optional): True = Use Excel Automation vs. OLE DB");
        }
    }
}
