using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Windows.Forms;
using System.Drawing;

namespace tabuleckaGenerator
{
    class TabuleckasGenerator
    {
        private Microsoft.Office.Interop.Excel.Application excelApp = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
        private int month;
        private int year;
        private int daysCount;

        public TabuleckasGenerator(int month, int year)
        {
            this.month = month;
            this.year = year;
            this.daysCount = GetDaysCount();

            CreateDocument();
        }

        public TabuleckasGenerator()
        {
            CreateDocument(false);
        }

        private void CreateDocument(bool setWorksheet = true)
        {
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add(1);

                if(setWorksheet)
                {
                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                    worksheet.Name = month + "_" + year;
                    //workSheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            excelApp = new Microsoft.Office.Interop.Excel.Application();
        }

        public void CreateHeaders(int row, int col, string text, Color backgroundColor)
        {
            worksheet.Cells[row, col] = text;
            var cellNumber = GetCellExcelNumber(row, col);
            var range = worksheet.get_Range(cellNumber, cellNumber);
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(backgroundColor);
        }

        public void CreateHeaders(int row, int col, string text)
        {
            worksheet.Cells[row, col] = text;
        }

        public void SaveDocument(string filePath)
        {
            workbook.SaveAs(filePath);
            workbook.Close(0);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        public void FillDateColumn(int startRow, int col)
        {
            for (int row = 0; row < daysCount; row++)
                worksheet.Cells[row + startRow, col] = new DateTime(year, month, row + 1);

            var range = worksheet.get_Range(GetCellExcelNumber(startRow, col), GetCellExcelNumber(startRow + daysCount - 1, col));
            range.NumberFormat = "dd.mm.yyyy";
        }

        public void SetCellsFormat(int startRow, int col, string format)
        {
            var range = worksheet.get_Range(GetCellExcelNumber(startRow, col), GetCellExcelNumber(startRow + daysCount - 1, col));
            range.NumberFormat = format;
        }

        public void FillDayColumn(int startRow, int col)
        {
            for (int row = 0; row < GetDaysCount(); row++)
                worksheet.Cells[row + startRow, col] = GetSkDay(new DateTime(year, month, row + 1).DayOfWeek);
        }

        public void SetSumCell(int startRow, int col, string name)
        {
            worksheet.Cells[daysCount + startRow, col] = name;
            var sumString = "=SUM(" + GetCellExcelNumber(startRow, col) + ":" + GetCellExcelNumber(startRow + daysCount - 1, col) + ")";
            worksheet.Cells[daysCount + startRow + 1, col].Formula = sumString;
        }

        public void FinalizeTable(int row, int startCol)
        {
            CreateHeaders(row, startCol, "-", Color.Red);
            worksheet.Cells[row + 1, startCol++].Formula = "=SUM(C2:F" + (daysCount + 1);

            CreateHeaders(row, startCol, "+", Color.Green);
            worksheet.Cells[row + 1, startCol++].Formula = "=SUM(G2:G" + (daysCount + 1);

            CreateHeaders(row, startCol, "=", Color.Blue);
            worksheet.Cells[row + 1, startCol++].Formula = "=-" + GetCellExcelNumber(row + 1, startCol - 3) + "+" + GetCellExcelNumber(row + 1, startCol - 2);

            worksheet.get_Range(GetCellExcelNumber(row, startCol - 3), GetCellExcelNumber(row, startCol - 1)).Cells.HorizontalAlignment =
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public void GenerateWholeYear(int year)
        {
            this.year = year;

            for (int i = 12; i >= 1; i--)
            {
                this.month = i;
                this.daysCount = GetDaysCount();

                //worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[i];
                if (i == 12)
                    worksheet = workbook.ActiveSheet;
                else
                    worksheet = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                worksheet.Name = GetMonthString();
                //workSheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;

                CreateHeaders(1, 1, "Deň", Color.LightBlue);
                CreateHeaders(1, 2, "Dátum", Color.LightBlue);

                CreateHeaders(1, 3, "Auto", Color.Red);
                SetCellsFormat(2, 3, "#,###,###.00€");
                SetSumCell(2, 3, "Auto");

                CreateHeaders(1, 4, "Obchod", Color.Red);
                SetCellsFormat(2, 4, "#,###,###.00€");
                SetSumCell(2, 4, "Obchod");

                CreateHeaders(1, 5, "Obedy", Color.Red);
                SetCellsFormat(2, 5, "#,###,###.00€");
                SetSumCell(2, 5, "Obedy");

                CreateHeaders(1, 6, "Iné", Color.Red);
                SetCellsFormat(2, 6, "#,###,###.00€");
                SetSumCell(2, 6, "Iné");

                CreateHeaders(1, 7, "Vklad", Color.Green);
                SetCellsFormat(2, 7, "#,###,###.00€");
                SetSumCell(2, 7, "Vklad");

                CreateHeaders(1, 8, "Výber", Color.LightBlue);
                SetCellsFormat(2, 8, "#,###,###.00€");
                SetSumCell(2, 8, "Výber");

                FillDayColumn(2, 1);
                FillDateColumn(2, 2);

                FinalizeTable(1, 10);
            }
        }

        private int GetDaysCount()
        {
            return DateTime.DaysInMonth(year, month);
        }

        private string GetSkDay(DayOfWeek day)
        {
            switch (day)
            {
                case DayOfWeek.Friday:
                    return "Piatok";
                case DayOfWeek.Monday:
                    return "Pondelok";
                case DayOfWeek.Saturday:
                    return "Sobota";
                case DayOfWeek.Sunday:
                    return "Nedeľa";
                case DayOfWeek.Thursday:
                    return "Štvrtok";
                case DayOfWeek.Tuesday:
                    return "Utorok";
                case DayOfWeek.Wednesday:
                    return "Streda";
                default:
                    return "Neznámy deň";
            }
        }

        private string GetMonthString()
        {
            switch (month)
            {
                case 1:
                    return "Január";
                case 2:
                    return "Február";
                case 3:
                    return "Marec";
                case 4:
                    return "Apríl";
                case 5:
                    return "Máj";
                case 6:
                    return "Jún";
                case 7:
                    return "Júl";
                case 8:
                    return "August";
                case 9:
                    return "September";
                case 10:
                    return "Október";
                case 11:
                    return "November";
                case 12:
                    return "December";
                default:
                    return "Neznámy mesiac";
            }
        }

        private string GetCellExcelNumber(int row, int col)
        {
            return NumberToUpperChar(col) + row;
        }

        private string NumberToUpperChar(int col)
        {
            return ((char)(64 + col)).ToString();
        }
    }
}
