using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace DemoProject
{
    public class ExcelWriter
    {
            /// <summary>
            /// Constructs a new instance.
            /// </summary>
            public ExcelWriter()
            {
                // Do not delete - a parameterless constructor is required!
            }


            public void Driver(int row, int col, string time, string sheetName)
            {

                string sDataFile = "Ranorex_Reports.xls";
                string sFilePath = Path.GetFullPath(sDataFile);

                string sOldvalue = "Automation\\bin\\Debug\\" + sDataFile;
                sFilePath = sFilePath.Replace(sOldvalue, "") + "PEPI_Performance\\ExecutionReport\\" + sDataFile;
                fnOpenExcel(sFilePath, sheetName);
                writeExcel(row, col, time);
                fnCloseExcel();
            }
            Excel.Application exlApp;
            Excel.Workbook exlWB;
            Excel.Sheets excelSheets;
            Excel.Worksheet exlWS;
            //Open Excel file
            public int fnOpenExcel(string sPath, string iSheet)
            {

                int functionReturnValue = 0;
                try
                {

                    exlApp = new Excel.Application(); //Microsoft.Office.Interop.Excel.Application();
                    exlApp.Visible = true;
                    exlWB = exlApp.Workbooks.Open(sPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // get all sheets in workbook
                    excelSheets = exlWB.Worksheets;

                    // get some sheet
                    //string currentSheet = "Cycle1";
                    exlWS = (Excel.Worksheet)excelSheets.get_Item(iSheet);
                    functionReturnValue = 0;
                }
                catch (Exception ex)
                {
                    functionReturnValue = -1;
                    Report.Error(ex.Message);
                }
                return functionReturnValue;
            }


            // Close the excel file and release objects.
            public int fnCloseExcel()
            {
                //exlWB.Close();

                try
                {
                    exlApp.ActiveWorkbook.Save();
                    exlApp.Quit();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(exlWS);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(exlWB);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(exlApp);

                    GC.GetTotalMemory(false);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.GetTotalMemory(true);
                }
                catch (Exception ex)
                {
                    Report.Error(ex.Message);
                }
                return 0;
            }

            public void writeExcel(int i, int j, string time)
            {
                Excel.Range exlRange = null;
                exlRange = (Excel.Range)exlWS.UsedRange;
                ((Excel.Range)exlRange.Cells[i, j]).Formula = time;

            }

        }
  }


