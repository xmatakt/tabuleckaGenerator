using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tabuleckaGenerator
{
    public partial class Form1 : Form
    {
        private TabuleckasGenerator generator = null;

        public Form1()
        {
            //this.generator = new TabuleckasGenerator(3, 2017);
            this.generator = new TabuleckasGenerator();

            InitializeComponent();

            //GenerateExcelDocument();
            this.generator.GenerateWholeYear(2017);
            generator.SaveDocument(@"C:\AATimo\daco.xlsx");
        }

        private void button1_Click(object sender, EventArgs e)
        {
    
        }

        private void GenerateExcelDocument()
        {
            generator.CreateHeaders(1, 1, "Deň", Color.LightBlue);
            generator.CreateHeaders(1, 2, "Dátum", Color.LightBlue);

            generator.CreateHeaders(1, 3, "Auto", Color.Red);
            generator.SetCellsFormat(2, 3, "#,###,###.00€");
            generator.SetSumCell(2, 3, "Auto");

            generator.CreateHeaders(1, 4, "Obchod", Color.Red);
            generator.SetCellsFormat(2, 4, "#,###,###.00€");
            generator.SetSumCell(2, 4, "Obchod");

            generator.CreateHeaders(1, 5, "Obedy", Color.Red);
            generator.SetCellsFormat(2, 5, "#,###,###.00€");
            generator.SetSumCell(2, 5, "Obedy");

            generator.CreateHeaders(1, 6, "Iné", Color.Red);
            generator.SetCellsFormat(2, 6, "#,###,###.00€");
            generator.SetSumCell(2, 6, "Iné");

            generator.CreateHeaders(1, 7, "Vklad", Color.Green);
            generator.SetCellsFormat(2, 7, "#,###,###.00€");
            generator.SetSumCell(2, 7, "Vklad");

            generator.CreateHeaders(1, 8, "Výber", Color.LightBlue);
            generator.SetCellsFormat(2, 8, "#,###,###.00€");
            generator.SetSumCell(2, 8, "Výber");

            generator.FillDayColumn(2, 1);
            generator.FillDateColumn(2, 2);

            generator.FinalizeTable(1, 10);

            generator.SaveDocument(@"C:\AATimo\daco.xlsx");
        }
    }
}
