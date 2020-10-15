using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace LekarTalon
{
    class Program
    {


        static async Task Main(string[] args)
        {

            Console.WriteLine("Идет печать талона");
            foreach (string arg in args)
            {
                string path = arg;
                using (FileStream fstream = File.OpenRead(path))
                {

                    byte[] array = new byte[fstream.Length];
                    await fstream.ReadAsync(array, 0, array.Length);

                    List<string> words = ((Encoding.GetEncoding("Windows-1251").GetString(array)).Split(new char[] { '\n' })).ToList();

                    words[0] = words[0].Remove(0, 8).Remove(14, 9);
                    words[1] = words[1].Remove(0, 23).Insert(0, "№ Карты: ");
                    words[2] = words[2].Remove(0, 6);
                    words[3] = words[3].Remove(0, 6);
                    words[4] = words[4].Remove(0, 12).Remove(5, 5).Insert(5, " ").Insert(0, "Приём в ");
                    words[5] = words[5].Remove(0, 18).Insert(0, "Кабинет № ");
                    words[6] = words[6].Remove(0, 4);

                    Application excelApp = new Application();
                    Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                    Worksheet _workSheet = excelApp.Sheets[1];
                    excelApp.Columns.ColumnWidth = 33;
                    (excelApp.Cells as Range).Font.Name = "Times New Roman";
                    (excelApp.Cells as Range).Font.Size = 14;
                    (excelApp.Cells as Range).WrapText = true;
                    (excelApp.Cells as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    (excelApp.Cells as Range).VerticalAlignment = XlVAlign.xlVAlignCenter;

                    excelApp.Sheets[1].PageSetup.LeftMargin = 0;
                    excelApp.Sheets[1].PageSetup.TopMargin = 0;
                    excelApp.Sheets[1].PageSetup.RightMargin = 0;

                    excelApp.Cells[1, 1] = words[0];
                    excelApp.Cells[2, 1] = words[1];
                    excelApp.Cells[3, 1] = words[2];
                    excelApp.Cells[4, 1] = words[3];
                    excelApp.Cells[5, 1] = words[4];
                    (excelApp.Cells[5, 1] as Range).Font.Bold = true;
                    (excelApp.Cells[5, 1] as Range).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    excelApp.Cells[6, 1] = words[5];
                    excelApp.Cells[7, 1] = words[6];
                    excelApp.Cells[8, 1] = words[7];
                    excelApp.Cells[9, 1] = ".";
                    excelApp.Cells[10, 1] = ".";
                    excelApp.Cells[11, 1] = ".";
                    excelApp.DisplayAlerts = false;
                    _workSheet.PrintOutEx(Type.Missing, Type.Missing, 1, Type.Missing, "prtalon", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //excelApp.Visible = true;

                    excelApp.Quit();
                    excelApp = null;
                    workbook = null;
                    _workSheet = null;

                    GC.Collect();

                }
            }
        }
    }
}
