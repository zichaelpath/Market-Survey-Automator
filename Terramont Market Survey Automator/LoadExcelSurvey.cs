using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
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
				CheckPathExists = true
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
			string filePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString();
			string fileName = "";
			Image property = new Bitmap(propertyImage.Image);
			
			SaveFileDialog saveFile = new SaveFileDialog(); 

			if (rdoFloorPlan.Checked)
			{
				string floorPlanFile = filePath + "\\Floor Plans\\";
				saveFile.InitialDirectory = floorPlanFile;
				saveFile.RestoreDirectory = true;
				
				saveFile.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp; *.png)|*.jpg; *.jpeg; *.gif; *.bmp; *.png";
				saveFile.DefaultExt = "*.jpg";
				fileName = cboProperties.Text.ToString() + "_FloorPlan" + Directory.GetFiles(floorPlanFile).Length.ToString();
			   saveFile.FileName = fileName;
				if (saveFile.ShowDialog() == DialogResult.OK)
				{
					System.IO.FileStream fileStream = (System.IO.FileStream)saveFile.OpenFile();
					property.Save(fileStream, System.Drawing.Imaging.ImageFormat.Jpeg);
					fileStream.Close();
				}
				string floorPlanFilePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\Floor Plans\\";
				int fileCount = Directory.GetFiles(floorPlanFilePath).Length;
				txtFloorPlansExist.Text = fileCount.ToString();
			}
			if (rdoGeneralImage.Checked)
			{
				string generalImageFile = filePath + "\\General Property Images\\";
				saveFile.InitialDirectory = generalImageFile;
				saveFile.RestoreDirectory = true;
				saveFile.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp; *.png)|*.jpg; *.jpeg; *.gif; *.bmp; *.png";
				saveFile.DefaultExt = "*.jpg";
				fileName = cboProperties.Text.ToString() + "_GeneralPropertyImages" + Directory.GetFiles(generalImageFile).Length.ToString();
				saveFile.FileName = fileName;   
				if (saveFile.ShowDialog() == DialogResult.OK)
				{
					System.IO.FileStream fileStream = (System.IO.FileStream)saveFile.OpenFile();
					property.Save(fileStream, System.Drawing.Imaging.ImageFormat.Jpeg);
					fileStream.Close();
				}
				string generalImagesFilePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\General Property Images\\";
				int fileCount = Directory.GetFiles(generalImagesFilePath).Length;
				txtPropertyImages.Text = fileCount.ToString();
			}

			saveFile.Dispose();
		}

		private void cboProperties_TextChanged(object sender, EventArgs e)
		{
			int fileCount = 0;
			string floorPlanFilePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\Floor Plans\\";
			fileCount = Directory.GetFiles(floorPlanFilePath).Length;
			txtFloorPlansExist.Text = fileCount.ToString();
			string generalImagesFilePath = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\General Property Images\\";
			fileCount = Directory.GetFiles(generalImagesFilePath).Length;
			txtPropertyImages.Text = fileCount.ToString(); 

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

		private void btnDeleteImage_Click(object sender, EventArgs e)
		{
			 if (rdoFloorPlan.Checked)
			{
				OpenFileDialog openFile = new OpenFileDialog();
				string floorPlanFilePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\Floor Plans\\";
				openFile.InitialDirectory = floorPlanFilePath;
				openFile.RestoreDirectory = false;
				
				if (openFile.ShowDialog() == DialogResult.OK)
				{
					File.Delete(openFile.FileName);
				}
				openFile.Dispose();
				int fileCount = 0;
				string floorPlan = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\Floor Plans\\";
				fileCount = Directory.GetFiles(floorPlan).Length;
				txtFloorPlansExist.Text = fileCount.ToString();
			}

			if (rdoGeneralImage.Checked)
			{
				OpenFileDialog openFile = new OpenFileDialog();
				string generalImagesFilePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\General Property Images\\";
				openFile.InitialDirectory = generalImagesFilePath;
				openFile.RestoreDirectory = false;

				if (openFile.ShowDialog() == DialogResult.OK)
				{
					File.Delete(openFile.FileName);
				}
				openFile.Dispose();
				int fileCount = 0;
				string generalImages = @"C:/Terramont Property Images/Properties/" + cboProperties.Text.ToString() + "\\General Property Images\\";
				fileCount = Directory.GetFiles(generalImages).Length;
				txtFloorPlansExist.Text = fileCount.ToString();
			}
		}
	}
}
