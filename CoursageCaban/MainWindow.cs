using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CoursageCaban
{
    public partial class MainWindow : Form
    {
        private Form closeForm;
        private Calculate calculate;
        private Office officeapp;

        public MainWindow(Form closeForm)
        {
            InitializeComponent();
            this.closeForm = closeForm;
            closeForm.Hide();
            officeapp = new Office();
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            calculate = new Calculate(0.05, 1, 0.1);
        }

        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            closeForm.Close();
        }

        private void calculateButton_Click(object sender, EventArgs e)
        {
            calculate.CalculateData();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            officeapp.OpenWord();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            officeapp.OpenExcel();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            closeForm.Close();
        }
    }
}
