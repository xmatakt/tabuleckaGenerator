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
            this.generator = new TabuleckasGenerator(2, 2017);

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            generator.FillDateColumn(1, 1);
            //generator.CreateHeaders(1,1,"HEllo WOrld");
            generator.SaveDocument(@"C:\AATimo\daco.xlsx");
        }
    }
}
