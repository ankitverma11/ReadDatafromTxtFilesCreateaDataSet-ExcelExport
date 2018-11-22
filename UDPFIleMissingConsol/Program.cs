using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;

namespace UDPFIleMissingConsol
{
    public class Program
    {
        static void Main(string[] args)
        {
            DataTable dtUpdatedata = Createdatable();
            string Folderpath = Convert.ToString(ConfigurationManager.AppSettings["UDPFileFolder"]);
            string DeviceSerialNumber = Convert.ToString(ConfigurationManager.AppSettings["DeviceSerialNumber"]);
            string SaveLocation = Convert.ToString(ConfigurationManager.AppSettings["SaveLocation"]);
            string fileName = SaveLocation + "UDPrecords";
            fileName = fileName + "_" + DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + ".xlsx";

            foreach (string sr in Directory.GetFiles(Folderpath).OrderBy(p => new FileInfo(p).LastWriteTime))
            {
                string fileNme = Path.GetFileName(sr);
               
                for (int i = 0; i <= File.ReadAllLines(sr, Encoding.UTF8).Count() - 1; i++)
                {
                    DataRow drUDP = dtUpdatedata.NewRow();
                    drUDP["FileName"] = fileNme; 
                    drUDP["DeviceSerialNumber"] = File.ReadAllLines(sr, Encoding.UTF8)[i].Split(',')[0];
                    if (DeviceSerialNumber.Equals(drUDP["DeviceSerialNumber"]))
                    {
                        drUDP["UniqueID"] = File.ReadAllLines(sr, Encoding.UTF8)[i].Split(',')[1];
                        drUDP["EventTime"] = Convert.ToDateTime(File.ReadAllLines(sr, Encoding.UTF8)[i].Split(',')[2]);
                        dtUpdatedata.Rows.Add(drUDP);
                    }
                }
                dtUpdatedata.AcceptChanges();
            }

            WriteDataTableToExcel(dtUpdatedata, "UDPFilesExcelReport", fileName);
        }

      private static DataTable Createdatable()
        {
            DataTable dt = new DataTable("UDP");
            try
            {
                dt.Columns.AddRange(new DataColumn[4] {
                    new DataColumn("DeviceSerialNumber",typeof(string)),
                    new DataColumn("UniqueID",typeof(string)),
                    new DataColumn("EventTime", typeof(DateTime)),
                    new DataColumn("Filename",typeof(string))
                });
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

      public static void WriteDataTableToExcel(System.Data.DataTable dataTable, string Worksheetname, string fileName)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = Worksheetname;

                // loop through each row and add values to our sheet
                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // on the first iteration we add the column headers
                        if (rowcount == 3)
                        {
                            excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
                        }

                        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        //for alternate rows
                        if (rowcount > 3)
                        {
                            if (i == dataTable.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                                }

                            }
                        }

                    }

                }
                // now we resize the columns
                excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
               // border.Weight = 2d;
                excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[2, dataTable.Columns.Count]];
                //now save the workbook and exit Excel
                excelworkBook.SaveAs(fileName); 
                excelworkBook.Close();
                excel.Quit();
             }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }

    }
}
