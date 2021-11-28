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
using System.IO;

namespace Terramont_Market_Survey_Automator
{
    public partial class LoadExcelSurvey : Form
    {
        ImageList propertyImages;
        List<ImageList> propertyImagesList = new List<ImageList>();
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
            string filePath = @"C:/Terramont Property Images/Properties";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
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
                excelDataGrid.Rows.Clear();
                excelDataGrid.Columns.Clear();

                for (int i = 1; i <= excelWorksheet.Columns.Count; i++)
                {
                    if (excelWorksheet.Cells[1, i].value == null)
                    {
                        break;      // BREAK LOOP.
                    }
                    else
                    {
                        DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                        col.HeaderText = excelWorksheet.Cells[1, i].value;
                        int colIndex = excelDataGrid.Columns.Add(col);        // ADD A NEW COLUMN.
                    }
                }

                for (int i = 2; i <= excelWorksheet.Rows.Count; i++)
                {
                    if (excelWorksheet.Cells[i, 1].value == null)
                    {
                        break;
                    }
                    else
                    {
                        string property = excelWorksheet.Cells[i, 1].value.ToString();
                        string[] row = new string[] { excelWorksheet.Cells[i, 1].value.ToString(),
                        excelWorksheet.Cells[i, 2].value.ToString(),
                        excelWorksheet.Cells[i, 3].value.ToString(),
                        excelWorksheet.Cells[i, 4].value.ToString(),
                        excelWorksheet.Cells[i, 5].value.ToString(),
                        excelWorksheet.Cells[i, 6].value.ToString(),
                        excelWorksheet.Cells[i, 7].value.ToString(),
                        excelWorksheet.Cells[i, 8].value.ToString(),
                        excelWorksheet.Cells[i, 9].value.ToString(),
                        excelWorksheet.Cells[i, 10].value.ToString(),
                        excelWorksheet.Cells[i, 11].value.ToString(),
                        excelWorksheet.Cells[i, 12].value.ToString(),
                        excelWorksheet.Cells[i, 13].value.ToString(),
                        excelWorksheet.Cells[i, 14].value.ToString()};

                        excelDataGrid.Rows.Add(row);
                        cboProperties.Items.Add(property);
                        filePath = @"C:/Terramont Property Images/Properties/" + property;
                        if (!Directory.Exists(filePath))
                        {
                            Directory.CreateDirectory(filePath);
                        }
                        string floorPlanFolder = filePath + "/Floor Plans";
                        if (!Directory.Exists(floorPlanFolder))
                        {
                            Directory.CreateDirectory(floorPlanFolder);
                        }
                        string propertyImagesFolder = filePath + "/General Property Images";
                        if (!Directory.Exists(propertyImagesFolder))
                        {
                            Directory.CreateDirectory(propertyImagesFolder);
                        }
                    }
                }

                excelWorkbook.Close();
                excelApp.Quit();


                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorksheet);
            }
        }

        private void LoadExcelSurvey_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void btnSaveImage_Click(object sender, EventArgs e)
        {
            string filePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString();
            string fileName = "";
            Image property;
            property = propertyImage.Image;
            SaveFileDialog saveFile = new SaveFileDialog(); 

            if (rdoFloorPlan.Checked)
            {
                string floorPlanFile = filePath + "/Floor Plans";
                saveFile.InitialDirectory = floorPlanFile;
                fileName = cboProperties.Text.ToString() + "_FloorPlan" + Directory.GetFiles(floorPlanFile).Length.ToString();
               saveFile.FileName = fileName;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    property.Save(fileName);
                }
               
            }
            if (rdoGeneralImage.Checked)
            {
                string generalImageFile = filePath + "/General Property Images";
                fileName = cboProperties.Text.ToString() + "_FloorPlan" + Directory.GetFiles(generalImageFile).Length.ToString();

                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    property.Save(fileName);
                }
            }
        }

        private void cboProperties_TextChanged(object sender, EventArgs e)
        {
            int fileCount = 0;
            string filePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString();
            fileCount = Directory.GetFiles(filePath).Length;
            txtNumOfProperties.Text = fileCount.ToString();

            if (fileCount == 0)
            {
                btnFirstImage.Enabled = false;
                btnLastImage.Enabled = false;
                btnNextImage.Enabled = false;
                btnPreviousImage.Enabled = false;
            }
            else if (fileCount > 2)
            {
                btnFirstImage.Enabled = true;
                btnPreviousImage.Enabled = true;
                btnNextImage.Enabled = true;
                btnLastImage.Enabled = true;
            }
        }
    }
}
