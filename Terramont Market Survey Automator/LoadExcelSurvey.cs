using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Terramont_Market_Survey_Automator
{
	public partial class LoadExcelSurvey : Form
	{

		Timer detailsCheck;
		List<Property> properties = new List<Property>();
		List<string> propertyFloorPlanDirectories = new List<string>();
		List<string> propertyGeneralImageDirectories = new List<string>();
		List<string> clientNeeds = new List<string>();
		List<Broker> brokerDetails = new List<Broker>();
		List<string> imagePaths = new List<string>();
		SurveyGenerator survey;
		public LoadExcelSurvey()
		{
			InitializeComponent();
			detailsCheck = new Timer();
			detailsCheck.Interval = 200;
			detailsCheck.Tick += new EventHandler(detailsCheck_click);
			detailsCheck.Start();
			Directory.GetCurrentDirectory().Contains("DefaultItems");
			imagePaths.Add(System.IO.Directory.GetCurrentDirectory() + "\\DefaultItems\\needsAnalysis.png");
			imagePaths.Add(System.IO.Directory.GetCurrentDirectory() + "\\DefaultItems\\reserachProcess.png");
			imagePaths.Add(System.IO.Directory.GetCurrentDirectory() + "\\DefaultItems\\availabilityOverview.png");
		
		}

		private void detailsCheck_click(object sender, EventArgs e)
		{
			string filePath = @"C:\Terramont Clients\" + txtClient.Text;
			

			if (txtArea.Text != "" && txtTerm.Text != "" && txtGrowth.Text != "" & txtRelocationObjectives.Text != "" && txtLocation.Text != ""
				&& txtBuildingType.Text != "" && txtParking.Text != "" && txtComments.Text != "" && txtClient.Text != "")
			{
				btnSaveNeeds.Enabled = true;
			}
			if (txtArea.Text == "" || txtTerm.Text == "" || txtGrowth.Text == "" || txtRelocationObjectives.Text == "" || txtLocation.Text == ""
				|| txtBuildingType.Text == "" || txtParking.Text == "" || txtComments.Text == "" || txtClient.Text == "")
			{
				btnSaveNeeds.Enabled = false;
			}

			if (clientNeeds.Count > 1 && properties.Count != 0 && propertyFloorPlanDirectories.Count > 0 &&
				propertyGeneralImageDirectories.Count > 0)
			{
				btnGenerateSurvey.Enabled = true;
			}
			else
			{
				btnGenerateSurvey.Enabled = false;
			}
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
						Property propertyInfo = new Property();
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

						propertyInfo.Address = excelWorksheet.Cells[i, 1].value.ToString();
						propertyInfo.Landlord = excelWorksheet.Cells[i, 2].value.ToString();
						propertyInfo.RentableArea = excelWorksheet.Cells[i, 3].value.ToString();
						propertyInfo.Term = excelWorksheet.Cells[i, 4].value.ToString();
						propertyInfo.Occupancy = excelWorksheet.Cells[i, 5].value.ToString();
						propertyInfo.Incentives = excelWorksheet.Cells[i, 6].value.ToString();
						propertyInfo.NetRent = excelWorksheet.Cells[i, 7].value.ToString();
						propertyInfo.OperationCosts = excelWorksheet.Cells[i, 8].value.ToString();
						propertyInfo.Taxes = excelWorksheet.Cells[i, 9].value.ToString();
						propertyInfo.Energy = excelWorksheet.Cells[i, 10].value.ToString();
						propertyInfo.TotalAdditionalRent = excelWorksheet.Cells[i, 11].value.ToString();
						propertyInfo.GrossRent = excelWorksheet.Cells[i, 12].value.ToString();
						propertyInfo.Parking = excelWorksheet.Cells[i, 13].value.ToString();
						propertyInfo.Comments = excelWorksheet.Cells[i, 14].value.ToString();
						properties.Add(propertyInfo);

						excelDataGrid.Rows.Add(row);
						cboProperties.Items.Add(property);
						filePath = @"C:/Terramont Property Images/Properties/" + property;
						if (!Directory.Exists(filePath))
						{
							Directory.CreateDirectory(filePath);
						}
						string floorPlanFolder = filePath + "\\Floor Plans";
						if (!Directory.Exists(floorPlanFolder))
						{
							Directory.CreateDirectory(floorPlanFolder);
						}
						string propertyImagesFolder = filePath + "\\General Property Images";
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

				txtFileList.Text += saveFile.FileName + "\r\n";
				propertyFloorPlanDirectories.Add(saveFile.FileName);
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

				
				txtFileList.Text += saveFile.FileName + "\r\n";
				propertyGeneralImageDirectories.Add(saveFile.FileName);
			
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

			string filePath;
			string[] files;
			if (rdoFloorPlan.Checked)
			{
				filePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\Floor Plans\\";
				files = Directory.GetFiles(filePath);
				
				if (txtFileList.Text != "")
				{
					txtFileList.Text = "";
				}

				

				for (int i = 0; i < files.Length; i++)
				{
					txtFileList.Text += files[i] + "\r\n";
					for (int j = 0; j < files.Length; j++)
					{
						if (propertyFloorPlanDirectories.Count == 0)
						{
							propertyFloorPlanDirectories.Add(files[i]);
						}
						if (files[i] != propertyFloorPlanDirectories[j])
						{
							propertyFloorPlanDirectories.Add(files[i]);
						}
						else
						{
							continue;
						}
					}
				}
			}
			if (rdoGeneralImage.Checked)
			{
				filePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\General Property Images\\";
				files = Directory.GetFiles(filePath);
				

				if (txtFileList.Text != "")
				{
					txtFileList.Text = "";
				}

				

				for (int i = 0; i < files.Length; i++)
				{
					
					txtFileList.Text += files[i] + "\r\n";
					for (int j = 0;j < files.Length; j++)
					{
						if (propertyGeneralImageDirectories.Count == 0)
						{
							propertyGeneralImageDirectories.Add(files[i]);
						}
						if (files[i] != propertyGeneralImageDirectories[j])
						{
							propertyGeneralImageDirectories.Add(files[i]);
						}
						else
						{
							continue;
						}
					}

				}
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

		

		private void rdoFloorPlan_CheckedChanged(object sender, EventArgs e)
		{
			string filePath;
			string[] files;
			if (cboProperties.Text != "-Select Properties To Add Images-" && rdoFloorPlan.Checked)
			{
				filePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\Floor Plans\\";
				files = Directory.GetFiles(filePath);
				string fileName = "";
				if (txtFileList.Text != "")
				{
					txtFileList.Text = "";
				}

				for (int i = 0; i < files.Length; i++)
				{
					fileName = Path.GetFileName(files[i]);
					txtFileList.Text += fileName + "\r\n";
					for (int j = 0; j < files.Length; j++)
					{
						if (propertyFloorPlanDirectories.Count == 0)
						{
							propertyFloorPlanDirectories.Add(files[i]);
						}
						if (files[i] != propertyFloorPlanDirectories[j])
						{
							propertyFloorPlanDirectories.Add(files[i]);
						}
						else
						{
							continue;
						}
					}

				}
			}
		}

		private void rdoGeneralImage_CheckedChanged(object sender, EventArgs e)
		{
			string filePath;
			string[] files;
			if (cboProperties.Text != "-Select Properties To Add Images-" && rdoGeneralImage.Checked)
			{
				filePath = @"C:\Terramont Property Images\Properties\" + cboProperties.Text.ToString() + "\\General Property Images\\";
				files = Directory.GetFiles(filePath);
				string fileName = "";

				if (txtFileList.Text != "")
				{
					txtFileList.Text = "";
				}

				for (int i = 0; i < files.Length; i++)
				{
					fileName = Path.GetFileName(files[i]);
					txtFileList.Text += fileName + "\r\n";
					for (int j = 0; j < files.Length; j++)
					{
						if (propertyGeneralImageDirectories.Count == 0)
						{
							propertyGeneralImageDirectories.Add(files[i]);
						}
						if (files[i] != propertyGeneralImageDirectories[j])
						{
							propertyGeneralImageDirectories.Add(files[i]);
						}
						else
						{
							continue;
						}
					}
				}
			}
		}

		private void btnSaveNeeds_Click(object sender, EventArgs e)
		{
			clientNeeds.Add(txtArea.Text);
			clientNeeds.Add(txtTerm.Text);
			clientNeeds.Add(txtGrowth.Text);
			clientNeeds.Add(txtRelocationObjectives.Text);
			clientNeeds.Add(txtLocation.Text);
			clientNeeds.Add(txtBuildingType.Text);
			clientNeeds.Add(txtParking.Text);
			clientNeeds.Add(txtComments.Text);
			clientNeeds.Add(txtClient.Text);
			string filePath = @"C:\Terramont Clients\" + txtClient.Text;
			if (!Directory.Exists(filePath))
			{
				Directory.CreateDirectory(filePath);
			}
			txtArea.Clear();
			txtTerm.Clear();
			txtGrowth.Clear();
			txtRelocationObjectives.Clear();
			txtLocation.Clear();
			txtBuildingType.Clear();
			txtParking.Clear();
			txtComments.Clear();
			txtClient.Clear();
		}

		private void btnGenerateSurvey_Click(object sender, EventArgs e)
		{
			bool rem = chkRem.Checked ? true : false;
			properties = properties.Distinct<Property>().ToList();
			propertyFloorPlanDirectories = propertyFloorPlanDirectories.Distinct<string>().ToList();
			propertyGeneralImageDirectories = propertyGeneralImageDirectories.Distinct<string>().ToList();
			brokerDetails = brokerDetails.Distinct<Broker>().ToList();
			if (cboLanguage.Text == "French")
			{
				survey = new SurveyGenerator(properties, propertyFloorPlanDirectories, propertyGeneralImageDirectories,
					 clientNeeds, imagePaths, true, rem);
				survey.CreateSurvey();
			}
			else if (cboLanguage.Text == "English")
			{
				survey = new SurveyGenerator(properties, propertyFloorPlanDirectories, propertyGeneralImageDirectories,
					 clientNeeds, imagePaths, false, rem);
				survey.CreateSurvey();
			}
			else
			{
				MessageBox.Show("Select a Language/Selectionnez une Langue!");
				cboLanguage.Focus();
			}
		}

		

		

		

		private void cboLanguage_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (cboLanguage.Text == "French")
			{
				lblArea.Text = "Superficie";
				lblTerm.Text = "Terme";
				lblGrowth.Text = "Croissance";
				lblRelocateObjectives.Text = "Objectifs";
				lblLocation.Text = "Emplacement";
				lblBuildingType.Text = "Type de Bâtiment";
				lblComments.Text = "Commentaires";


				lblFloorPlans.Text = "Plans d'étage";
				lblPropertyImages.Text = "Photos de Propriété";
				rdoFloorPlan.Text = "Plans d'étage";
				rdoGeneralImage.Text = "Photos de Propriété";
				lblPropertyList.Text = "Liste des Images de Propriété";

				btnSelectImage.Text = "Sélectionner";
				btnSaveImage.Text = "Sauver";
				btnDeleteImage.Text = "Effacer";
			}
			if (cboLanguage.Text == "English")
			{
				lblArea.Text = "Area";
				lblTerm.Text = "Term";
				lblGrowth.Text = "Growth";
				lblRelocateObjectives.Text = "Relocation Objectives";
				lblLocation.Text = "Location";
				lblBuildingType.Text = "Building Type";
				lblComments.Text = "Comments";


				lblFloorPlans.Text = "Floor Plans";
				lblPropertyImages.Text = "Property Images";
				rdoFloorPlan.Text = "Floor Plan";
				rdoGeneralImage.Text = "General Image";
				lblPropertyList.Text = "List of Property Images";

				btnSelectImage.Text = "Select";
				btnSaveImage.Text = "Save";
				btnDeleteImage.Text = "Delete";
			}
		}

		
	}
}
