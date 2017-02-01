using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Windows.Forms;

namespace tabuleckaGenerator
{
    class TabuleckasGenerator
    {
        private Microsoft.Office.Interop.Excel.Application excelApp = null;
        private Microsoft.Office.Interop.Excel.Workbook workbook = null;
        private Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
        private int month;
        private int year;

        public TabuleckasGenerator(int month, int year)
        {
            this.month = month;
            this.year = year;

            CreateDocument();
        }

        private void CreateDocument()
        {
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add(1);

                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                worksheet.Name = month + "_" + year;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            excelApp = new Microsoft.Office.Interop.Excel.Application();
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
    }
}
