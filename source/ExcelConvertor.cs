using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace NebulaRnD.Utils.NebulaXConvert
{
    internal class ExcelConvertor : IDisposable
    {
        private const string OPTION_ERROR = "e";
        private readonly string fromPath;
        private readonly string options;
        private readonly string toPath;
        private Application app = null;
        private bool displayAlerts;
        private Workbook wb = null;

        public ExcelConvertor(string infile, string outfile, string options)
        {
            fromPath = infile;
            toPath = outfile;
            this.options = options;
            Convert();
        }

        public void Dispose()
        {
            Wrapup();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void Wrapup()
        {
            if (wb != null)
            {
                try
                {
                    wb.Close(SaveChanges: false);
                    app.DisplayAlerts = displayAlerts; // reset to whatever it was before
                }
                finally
                {
                    wb = null;
                }
            }
            if (app != null)
            {
                try
                {
                    app.Quit();
                }
                finally
                {
                    app = null;
                }
            }
        }

        private void Convert()
        {
            bool OK = false;
            try
            {
                if (OpenWB())
                {
                    if (SaveWB())
                    {
                        OK = true;
                    }
                }
            }
            catch
            {
            }
            if (options.Contains(OPTION_ERROR))
            {
                Console.WriteLine(string.Format("RESULT:{0}{1}", Environment.NewLine, OK ? "yes" : "no"));
            }
            else
            {
                Console.WriteLine(OK ? "yes" : "no");
            }
            try
            {
                Wrapup();
            }
            catch
            {
            }
        }

        private bool OpenWB()
        {
            Exception hold = null;
            try
            {
                app = new Application();
                wb = app.Workbooks.Open(fromPath.Contains(@"\") ? fromPath : string.Format("{0}{1}{2}", Environment.CurrentDirectory, @"\", fromPath));
                // on successful open turn off alert displays
                displayAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                hold = ex;
                if (options.Contains(OPTION_ERROR))
                {
                    Console.WriteLine(string.Format("ERROR: Opening:{0}{1}", Environment.NewLine, ex.Message));
                }
            }
            if (hold != null)
            {
                return false;
            }
            return true;
        }

        private bool SaveWB()
        {
            try
            {
                XlFileFormat format;
                if (Path.GetExtension(toPath).ToLower() == ".xls")
                {
                    format = XlFileFormat.xlWorkbookNormal;
                }
                else if (Path.GetExtension(toPath).ToLower() == ".xlsx")
                {
                    format = XlFileFormat.xlWorkbookDefault;
                }
                else
                {
                    throw new Exception("Extension must be XLS for Office 2003 or XLSX for 2007+");
                }
                wb.SaveAs(
                    FileFormat: format,
                    Filename: toPath.Contains(@"\") ? toPath : string.Format("{0}{1}{2}", Environment.CurrentDirectory, @"\", toPath),
                    AccessMode: XlSaveAsAccessMode.xlExclusive,
                    ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
            }
            catch (Exception ex)
            {
                if (options.Contains(OPTION_ERROR))
                {
                    Console.WriteLine(string.Format("ERROR: Saving:{0}{1}", Environment.NewLine, ex.Message));
                }
                return false;
            }
            return true;
        }
    }
}