using Microsoft.Office.Interop.Excel;
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

        private double h = 2, d = 0.1, R = 2;

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

            using (OpenGL game = new OpenGL(800, 600, "LearnOpenTK"))
            {
                //Run takes a double, which is how many frames per second it should strive to reach.
                //You can leave that out and it'll just update as fast as the hardware will allow it.
                game.Run(60.0);
            }
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

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "км.":
                {
                    h_textBox.Text = $"{h / 1000}";
                    break;
                }
                case "м.":
                {
                    h_textBox.Text = $"{h}";
                    break;
                }
                case "дм.":
                {
                    h_textBox.Text = $"{h * 10}";
                    break;
                }
            }
        }
    }
}
