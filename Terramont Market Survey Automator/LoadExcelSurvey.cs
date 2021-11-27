using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Terramont_Market_Survey_Automator
{
    public partial class LoadExcelSurvey : Form
    {
        public LoadExcelSurvey()
        {
            InitializeComponent();
        }

        private void btnSelectImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:/",
                Title = "Browse Images",
                CheckFileExists = true,
                CheckPathExists = true,
                RestoreDirectory = true,
                Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                propertyImage.Image = new Bitmap(openFileDialog.FileName);
            }
        }

        private void LoadExcelSurvey_Load(object sender, EventArgs e)
        {
            propertyImage.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        private void btnLoadSurvey_Click(object sender, EventArgs e)
        {
            string file = "";
            OpenFileDialog excelDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:/",
                Title = "Load Survey",
                CheckFileExists = true,
                CheckPathExists = true,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog.FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);
                Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        excelDataGrid.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                    }
                    break;
                }
            }

        }
    }
}
