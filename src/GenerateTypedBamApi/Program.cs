// BizTalk Typed BAM API Generator
// Copyright (C) 2008-Present Thomas F. Abraham. All Rights Reserved.
// Copyright (c) 2007 Darren Jefford. All Rights Reserved.
// Licensed under the MIT License. See License.txt in the project root.

using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace GenerateTypedBAMAPI
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length < 4)
            {
                PrintUsage();
                return -1;
            }

            string xlsFileName = args[0];
            string outputFileName = args[1];
            string eventStreamType = args[2].ToLower();
            string targetNamespace = args[3];
            string xsltPath = null;

            if (((eventStreamType != "direct") && (eventStreamType != "buffered")) && (eventStreamType != "orchestration"))
            {
                PrintUsage();
                PrintError("Unknown EventStream type, use Direct, Buffered or Orchestration");
                return -1;
            }

            if (args.Length == 5)
            {
                if (!string.IsNullOrEmpty(args[4].Trim()))
                {
                    xsltPath = args[4].Trim();

                    if (!File.Exists(xsltPath))
                    {
                        PrintUsage();
                        PrintError("Cannot find the specified XSLT file '" + xsltPath + "'.");
                    }
                }
            }

            char ch = eventStreamType[0];
            eventStreamType = ch.ToString().ToUpper() + eventStreamType.Substring(1);
            PrintHeader();

            try
            {
                if (!string.IsNullOrEmpty(xsltPath))
                {
                    Console.WriteLine("Using custom XSLT file " + Path.GetFileName(xsltPath));
                    Console.WriteLine();
                }

                Console.Write("Retrieving BAM Definition from the Excel Spreadsheet... ");
                string bAMDefinitionXML = Shared.BamDefinitionXmlExporter.GetBamDefinitionXml(xlsFileName, false);
                Console.WriteLine("Success");

                Console.Write("Generating typed C# API from the BAM XML... ");
                GenerateBAMAPIFromXML(bAMDefinitionXML, outputFileName, eventStreamType, targetNamespace, xsltPath);
                Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                PrintError(ex.Message);
                return -1;
            }

            return 0;
        }

        private static void PrintError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("ERROR:" + message);
            Console.ResetColor();
        }

        private static void PrintHeader()
        {
            Console.ForegroundColor = ConsoleColor.White;

            Version assemblyVersion = System.Reflection.Assembly.GetEntryAssembly().GetName().Version;
            Console.WriteLine(
                "BizTalk BAM API Generator "
                + assemblyVersion.Major + "." + assemblyVersion.Minor + "." + assemblyVersion.Build);
            Console.WriteLine("[https://github.com/tfabraham/BizTalkTypedBAMAPIGenerator]");
            Console.WriteLine("Copyright (C) 2007 Darren Jefford and 2008 Thomas F. Abraham");
            Console.WriteLine();

            Console.ResetColor();
        }

        private static void PrintUsage()
        {
            PrintHeader();
            Console.WriteLine("Generates a typed C# BAM API from an Excel BAM Definition spreadsheet");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("GenerateTypedBAMAPI <ExcelFile> <CodeFile> <Direct|Buffered|Orchestration> <.NET Namespace> [XSLTFile]");
            Console.WriteLine(@"e.g: GenerateTypedBAMAPI C:\BAMDef.xls C:\BAMAPI.cs Buffered BizTalk.BAM");
            Console.WriteLine();
            Console.WriteLine("[XSLTFile]: Optional path to XSLT file that creates BAM API code from BAM XML");
        }

        private static void GenerateBAMAPIFromXML(
            string bamDefinitionXml, string outputFileName, string eventStreamType, string targetNamespace, string xsltPath)
        {
            XsltSettings settings = new XsltSettings(false, true);
            
            // Read the XSLT used to create the C# from the BAM def XML
            StringReader xsltReader = null;

            if (string.IsNullOrEmpty(xsltPath))
            {
                xsltReader = new StringReader(Properties.Resources.TypedApi);
            }
            else
            {
                xsltReader = new StringReader(File.ReadAllText(xsltPath));
            }

            XPathDocument stylesheet = new XPathDocument(xsltReader);

            XslCompiledTransform transform = new XslCompiledTransform();
            transform.Load(stylesheet, settings, null);

            // Read the BAM def XML
            StringReader textReader = new StringReader(bamDefinitionXml);
            XPathDocument document = new XPathDocument(textReader);

            Version assemblyVersion = System.Reflection.Assembly.GetEntryAssembly().GetName().Version;

            XsltArgumentList arguments = new XsltArgumentList();
            arguments.AddParam("EventStreamType", "", eventStreamType);
            arguments.AddParam("TargetNamespace", "", targetNamespace);
            arguments.AddParam("ToolVersion", "", assemblyVersion.Major + "." + assemblyVersion.Minor + "." + assemblyVersion.Build);

            using (StreamWriter writer = new StreamWriter(outputFileName))
            {
                transform.Transform(document, arguments, writer);
            }
        }
    }
}
