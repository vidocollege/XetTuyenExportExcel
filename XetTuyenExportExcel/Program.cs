using Newtonsoft.Json;
using System;
using XetTuyenExportExcel.Respository;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XetTuyenExportExcel.Models;
using System.IO;
using System.Reflection;

namespace XetTuyenExportExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            XetTuyenResposity XetTuyenResposity = new XetTuyenResposity();
            var dataBody = XetTuyenResposity.GetXetTuyen();
            dt = (DataTable)JsonConvert.DeserializeObject(dataBody.Result, typeof(DataTable));
            Save_data_table_to_excel(dt);
            Console.ReadKey();
        }
        public static void Save_data_table_to_excel(DataTable dt)
        {
            int intHeaderLength = 3;
            int intColumn = 0;
            int intRow = 0;
            string Work_sheet_name = "DanhSach";
            string Report_Type = "Details";
            System.Reflection.Missing Default = System.Reflection.Missing.Value;

            //create the excel file
            //string FilePath = @"\\Excel" + DateTime.Now.ToString().Replace(":", "_" + ".xlsx");
            string FolderExcel = Directory.GetCurrentDirectory() + "\\ExcelTemplete\\";
            string FilePath = FolderExcel + @"FormXetTuyen.xls";

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkbook;
            Microsoft.Office.Interop.Excel.Worksheet excelsheet;
            Microsoft.Office.Interop.Excel.Range excelCellRange;
            try
            {

                //start the application
                excel = new Microsoft.Office.Interop.Excel.Application();
                if (excel == null)
                {
                    Console.WriteLine("getting null values");
                }

                //for making excel visiable
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // creation a new work book
                excelworkbook = excel.Workbooks.Open(FilePath, Type.Missing, false, Type.Missing, Type.Missing,
            Type.Missing, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                //excelsheet
                excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkbook.ActiveSheet;
                excelsheet.Name = Work_sheet_name;
                int rowcount = 7;
                int stt = 1;
                foreach (DataRow datarow in dt.Rows)
                {
                    excelsheet.Cells[rowcount - 1, 1] = stt;
                    excelsheet.Cells[rowcount - 1, 1].EntireRow.Font.Bold = false;
                    excelsheet.Cells[rowcount - 1, 2] = datarow[0].ToString();
                    excelsheet.Cells[rowcount - 1, 3] = datarow[1].ToString();
                    excelsheet.Cells[rowcount - 1, 4] = "'" + datarow[2].ToString();
                    excelsheet.Cells[rowcount - 1, 5] = datarow[3].ToString();
                    excelsheet.Cells[rowcount - 1, 6] = datarow[4].ToString();
                    excelsheet.Cells[rowcount - 1, 7] = datarow[5].ToString();
                    excelsheet.Cells[rowcount - 1, 8] = datarow[6].ToString();
                    excelsheet.Cells[rowcount - 1, 9] = datarow[7].ToString();
                    excelsheet.Cells[rowcount - 1, 10] = datarow[8].ToString();
                    excelsheet.Cells[rowcount - 1, 11] = datarow[9].ToString();
                    excelsheet.Rows[rowcount].Insert();
                    stt++;
                    rowcount += 1;
                    excelCellRange = excelsheet.Range[excelsheet.Cells[rowcount, 2], excelsheet.Cells[rowcount, dt.Columns.Count]];
                    excelCellRange.EntireColumn.AutoFit();
                    if (stt > dt.Rows.Count)
                    {
                        excelsheet.Cells[rowcount + 3, 5] = dt.Rows.Count;
                    }
                }
               
                //now save the work book and exit the ecel
                var temp = "Danh sách đăng kí xét tuyển";
                var FolderExcelSave = Directory.GetCurrentDirectory() + "\\Excel\\";
                excelworkbook.SaveAs(FolderExcelSave + temp);
                excelworkbook.Close();
                excel.Quit();
                Console.WriteLine($"Tao file excel {temp} thanh cong");
                excelworkbook = excel.Workbooks.Open(FolderExcelSave+temp);
                excel.Visible = true;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);

            }
            finally
            {
                excelsheet = null;
                excelCellRange = null;
                excelworkbook = null;
            }

        }

    }
}
