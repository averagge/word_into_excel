using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;
using OfficeOpenXml;
using System.IO;

namespace word_into_excel
{
    public class WordClass
    {
        private List<Table> subTables = new List<Table>();
        private List<Crime> crimes = new List<Crime>();
        public void CreateSubTables(string filename)
        {

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, false))
            {
                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                Table subTable = new Table();
                foreach (TableRow row in table.Elements<TableRow>())
                {
                    if (row.Elements<TableCell>().All(cell => string.IsNullOrEmpty(cell.InnerText.Trim())))
                    {
                        subTables.Add(subTable);
                        subTable = new Table();
                    }
                    else
                    {
                        subTable.AppendChild(row.CloneNode(true));
                    }
                }
                subTables.Add(subTable);

            }
            
        }
        private int FindDescriptionStartPosition(Table table)
        {
            int rowindex = 0;
            foreach (TableRow row in table.Elements<TableRow>())
            {
                if (row.Elements<TableCell>().ElementAt(0).InnerText.Contains("Время прибытия наряда:"))
                {
                    return rowindex + 1;
                }
                rowindex++;
            }
            return rowindex;
        }
        public void GetAllData(string filename)
        {
            CreateSubTables(filename);
            string date = DateTime.Now.ToString("dd/MM/yyyy");
            foreach (Table table in subTables)
            {
                try
                {
                    TableCell titlecell = table.Elements<TableRow>().ElementAt(0).Elements<TableCell>().ElementAt(0);
                    string title = titlecell.InnerText.Split('.')[1].TrimStart();
                    TableCell departmentcell = table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(0);
                    string department = departmentcell.InnerText;
                    TableCell kuspcell = table.Elements<TableRow>().ElementAt(3).Elements<TableCell>().ElementAt(0);
                    string kusp = kuspcell.InnerText.Split(' ')[1];
                    string description = "";
                    int tablerow = table.Elements<TableRow>().Count() - 2;
                    int rowindex = this.FindDescriptionStartPosition(table);
                    for (int i = rowindex; i < tablerow; i++)
                    {
                        TableCell cell = table.Elements<TableRow>().ElementAt(i).Elements<TableCell>().ElementAt(0);
                        description += cell.InnerText + '\n';
                    }
                    description = description.Trim('\n');
                    TableCell infaboutinitiationcell = table.Elements<TableRow>().ElementAt(tablerow).Elements<TableCell>().ElementAt(0);
                    string infaboutinitiation = "";
                    try
                    {
                        infaboutinitiation = infaboutinitiationcell.InnerText.Split(':')[1].TrimStart();
                    }
                    catch { }
                    Crime crime = new Crime(date, title, department, kusp, description, infaboutinitiation);
                    crimes.Add(crime);

                }
                catch { }

            }
            subTables.Clear();

        }

        public void WriteIntoExcel(string filename)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filename)))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];

                int lastUsedRow = worksheet.Dimension.End.Row;

                // Определяем, куда добавить новую строку (например, следующую за последней используемой строкой)
                int newRow = lastUsedRow + 1;

                // Заполняем новую строку данными
                int counter = 0;
                foreach (Crime crime in crimes)
                {
                    worksheet.Cells[newRow + counter, 1].Value = crime._date;
                    worksheet.Cells[newRow + counter, 2].Value = crime._title;
                    worksheet.Cells[newRow + counter, 3].Value = crime._department;
                    worksheet.Cells[newRow + counter, 4].Value = crime._kusp;
                    worksheet.Cells[newRow + counter, 5].Value = crime._description;
                    worksheet.Cells[newRow + counter, 6].Value = crime._infaboutinitiation;
                    counter++;
                }
                // Сохраняем изменения
                excelPackage.Save();
            }

            /*Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filename);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int lastUsedRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            int newRow = lastUsedRow + 1;

            int counter = 0;
            foreach (Crime crime in crimes)
            {
                worksheet.Cells[newRow + counter, 1].Value = crime._date;
                worksheet.Cells[newRow + counter, 2].Value = crime._title;
                worksheet.Cells[newRow + counter, 3].Value = crime._department;
                worksheet.Cells[newRow + counter, 4].Value = crime._kusp;
                worksheet.Cells[newRow + counter, 5].Value = crime._description;
                worksheet.Cells[newRow + counter, 6].Value = crime._infaboutinitiation;
                counter++;
            }
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            crimes.Clear();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);*/

            MessageBox.Show("Готово");

        }
    }
}
