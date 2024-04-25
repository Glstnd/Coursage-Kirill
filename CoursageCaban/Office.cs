using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace CoursageCaban
{
    internal class Office
    {
        private Word.Application wordApp;
        private Excel.Application excelApp;

        public Office()
        {
            wordApp = new Word.Application();
            excelApp = new Excel.Application();
        }

        public void OpenWord()
        {
            wordApp.Visible = true;
        }

        public void OpenExcel()
        {
            excelApp.Visible = true;
        }
    }
}
