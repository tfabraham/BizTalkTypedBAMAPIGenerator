// BizTalk Typed BAM API Generator
// Copyright (C) 2008-Present Thomas F. Abraham. All Rights Reserved.
// Copyright (c) 2007 Darren Jefford. All Rights Reserved.
// Licensed under the MIT License. See License.txt in the project root.

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Data.OleDb;
using System.Data;

namespace Shared
{
    /// <summary>
    /// Extracts the XML for a BizTalk BAM definition from a binary Excel XLS file.
    /// </summary>
    internal static class BamDefinitionXmlExporter
    {
        private static object objMissing = Missing.Value;

        public static string GetBamDefinitionXml(string xlsFileName, bool useAutomation)
        {
            if (!File.Exists(xlsFileName))
            {
                throw new ArgumentException("File '" + xlsFileName + "' does not exist or is unavailable.");
            }

            if (useAutomation)
            {
                return GetBamDefinitionXmlAutomation(xlsFileName);
            }
            else
            {
                return GetBamDefinitionXmlDirect(xlsFileName);
            }
        }

        private static string GetBamDefinitionXmlDirect(string xlsFileName)
        {
            DataSet ds = new DataSet();
            OleDbConnection conn = null;

            try
            {
                conn = GetOleDbConnection(xlsFileName);

                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [BamXmlHiddenSheet$]", conn);

                try
                {
                    da.Fill(ds);
                }
                catch (OleDbException ex)
                {
                    if (ex.Errors.Count > 0 && ex.Errors[0].NativeError == -537199594)
                    {
                        throw new ArgumentException("ERROR: Could not find hidden BAM worksheet BamXmlHiddenSheet.", ex);
                    }

                    throw;
                }
            }
            finally
            {
                if (conn != null)
                {
                    conn.Dispose();
                }
            }

            // We should have gotten back a single DataTable with one row containing one or more columns
            // with a non-DBNull value.
            if (ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0 || ds.Tables[0].Rows[0].IsNull(0))
            {
                throw new ArgumentException(
                    "ERROR: Could not find hidden BAM worksheet or found no BAM XML on the worksheet. Expected to find BAM XML at cell BamXmlHiddenSheet!A1.");
            }

            // Build the complete XML by appending all cell values across the column.
            StringBuilder sb = new StringBuilder();

            DataRow row = ds.Tables[0].Rows[0];

            int columnCount = ds.Tables[0].Columns.Count;

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                if (!row.IsNull(columnIndex))
                {
                    sb.Append(row[columnIndex].ToString());
                }
                else
                {
                    break;
                }
            }
            
            return sb.ToString();
        }

        private static OleDbConnection GetOleDbConnection(string xlsFileName)
        {
            OleDbConnection conn = null;
            string connectionString = null;

            // Try Data Connnectivity Components 2007
            connectionString =
                string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=NO;MAXSCANROWS=1\"", xlsFileName);

            conn = TryOpenOleDbConnection(connectionString);

            if (conn != null)
            {
                return conn;
            }

            // Try Access Database Engine 2010 Redistributable
            connectionString =
                string.Format("Provider=Microsoft.ACE.OLEDB.14.0;Data Source={0};Extended Properties=\"Excel 14.0;HDR=NO;MAXSCANROWS=1\"", xlsFileName);

            conn = TryOpenOleDbConnection(connectionString);

            if (conn != null)
            {
                return conn;
            }

            // Try old Jet driver if XLS
            if (string.Compare(Path.GetExtension(xlsFileName), ".xls", StringComparison.OrdinalIgnoreCase) == 0)
            {
                connectionString =
                    string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=NO;MAXSCANROWS=1\"", xlsFileName);

                conn = TryOpenOleDbConnection(connectionString);

                if (conn != null)
                {
                    return conn;
                }
            }

            throw new Exception(
                "ERROR: Could not open an OLE DB connection to the Excel file. " +
                "Export from XLSX requires installation of Microsoft Office 2007 or Access Database Engine 2010 Redistributable or JET 4.0 for XLS.");
        }

        private static OleDbConnection TryOpenOleDbConnection(string connectionString)
        {
            OleDbConnection conn = new OleDbConnection(connectionString);

            try
            {
                conn.Open();
            }
            catch (Exception)
            {
                return null;
            }

            return conn;
        }

        private static string GetBamDefinitionXmlAutomation(string xlsFileName)
        {
            // This is needed to avoid issues with other cultures (see http://support.microsoft.com/kb/320369).
            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US");

            object excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            Type excelAppType = excelApp.GetType();

            excelAppType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false }, ci);
            excelAppType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excelApp, new object[] { false }, ci);
            object workbooks = excelAppType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null, ci);

            string fullXLSPath = Path.GetFullPath(xlsFileName);
            object[] args = new object[] { fullXLSPath, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing };
            object workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, args, ci);

            object worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null, ci);
            object[] objArray2 = new object[] { "BamXmlHiddenSheet" };
            object bamWorksheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, objArray2, ci);
            object cells = bamWorksheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, bamWorksheet, null, ci);

            StringBuilder builder = new StringBuilder();
            int num = 1;
            object[] objArray3 = new object[] { 1, num++ };
            object cell = cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, objArray3, ci);

            string str = (string)cell.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, cell, null, ci);

            for (;
                (str != null) && (str.Length > 0);
                str = (string)cell.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, cell, null, ci))
            {
                builder.Append(str);
                objArray3[1] = num++;
                cell = cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, objArray3, ci);
            }

            excelAppType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null, ci);
            excelApp = null;

            return builder.ToString();
        }
    }
}
