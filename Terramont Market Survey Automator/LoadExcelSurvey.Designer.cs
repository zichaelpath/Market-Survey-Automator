namespace Terramont_Market_Survey_Automator
{
	partial class LoadExcelSurvey
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadExcelSurvey));
			this.btnLoadSurvey = new System.Windows.Forms.Button();
			this.documentPreview = new System.Windows.Forms.TabControl();
			this.tabAvailabilitiesOverview = new System.Windows.Forms.TabPage();
			this.excelDataGrid = new System.Windows.Forms.DataGridView();
			this.tabAvailabilitiesMap = new System.Windows.Forms.TabPage();
			this.lblPropertyList = new System.Windows.Forms.Label();
			this.txtFileList = new System.Windows.Forms.TextBox();
			this.rdoGeneralImage = new System.Windows.Forms.RadioButton();
			this.rdoFloorPlan = new System.Windows.Forms.RadioButton();
			this.btnDeleteImage = new System.Windows.Forms.Button();
			this.btnSaveImage = new System.Windows.Forms.Button();
			this.btnSelectImage = new System.Windows.Forms.Button();
			this.propertyImage = new System.Windows.Forms.PictureBox();
			this.txtPropertyImages = new System.Windows.Forms.TextBox();
			this.txtFloorPlansExist = new System.Windows.Forms.TextBox();
			this.lblPropertyImages = new System.Windows.Forms.Label();
			this.lblFloorPlans = new System.Windows.Forms.Label();
			this.cboProperties = new System.Windows.Forms.ComboBox();
			this.tabNeeds = new System.Windows.Forms.TabPage();
			this.btnSaveNeeds = new System.Windows.Forms.Button();
			this.txtClient = new System.Windows.Forms.TextBox();
			this.txtParking = new System.Windows.Forms.TextBox();
			this.txtComments = new System.Windows.Forms.TextBox();
			this.txtBuildingType = new System.Windows.Forms.TextBox();
			this.txtLocation = new System.Windows.Forms.TextBox();
			this.txtRelocationObjectives = new System.Windows.Forms.TextBox();
			this.txtGrowth = new System.Windows.Forms.TextBox();
			this.txtTerm = new System.Windows.Forms.TextBox();
			this.txtArea = new System.Windows.Forms.TextBox();
			this.lblClientName = new System.Windows.Forms.Label();
			this.lblComments = new System.Windows.Forms.Label();
			this.lblParking = new System.Windows.Forms.Label();
			this.lblBuildingType = new System.Windows.Forms.Label();
			this.lblLocation = new System.Windows.Forms.Label();
			this.lblRelocateObjectives = new System.Windows.Forms.Label();
			this.lblGrowth = new System.Windows.Forms.Label();
			this.lblTerm = new System.Windows.Forms.Label();
			this.lblArea = new System.Windows.Forms.Label();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btnGenerateSurvey = new System.Windows.Forms.Button();
			this.bkgProcessFile = new System.ComponentModel.BackgroundWorker();
			this.cboLanguage = new System.Windows.Forms.ComboBox();
			this.chkRem = new System.Windows.Forms.CheckBox();
			this.documentPreview.SuspendLayout();
			this.tabAvailabilitiesOverview.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.excelDataGrid)).BeginInit();
			this.tabAvailabilitiesMap.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.propertyImage)).BeginInit();
			this.tabNeeds.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnLoadSurvey
			// 
			this.btnLoadSurvey.Location = new System.Drawing.Point(182, 48);
			this.btnLoadSurvey.Margin = new System.Windows.Forms.Padding(6);
			this.btnLoadSurvey.Name = "btnLoadSurvey";
			this.btnLoadSurvey.Size = new System.Drawing.Size(240, 71);
			this.btnLoadSurvey.TabIndex = 1;
			this.btnLoadSurvey.Text = "Load Excel Survey";
			this.btnLoadSurvey.UseVisualStyleBackColor = true;
			this.btnLoadSurvey.Click += new System.EventHandler(this.btnLoadSurvey_Click);
			// 
			// documentPreview
			// 
			this.documentPreview.Controls.Add(this.tabAvailabilitiesOverview);
			this.documentPreview.Controls.Add(this.tabAvailabilitiesMap);
			this.documentPreview.Controls.Add(this.tabNeeds);
			this.documentPreview.Location = new System.Drawing.Point(70, 177);
			this.documentPreview.Margin = new System.Windows.Forms.Padding(6);
			this.documentPreview.Name = "documentPreview";
			this.documentPreview.SelectedIndex = 0;
			this.documentPreview.Size = new System.Drawing.Size(1708, 1287);
			this.documentPreview.TabIndex = 2;
			// 
			// tabAvailabilitiesOverview
			// 
			this.tabAvailabilitiesOverview.Controls.Add(this.excelDataGrid);
			this.tabAvailabilitiesOverview.Location = new System.Drawing.Point(8, 39);
			this.tabAvailabilitiesOverview.Margin = new System.Windows.Forms.Padding(6);
			this.tabAvailabilitiesOverview.Name = "tabAvailabilitiesOverview";
			this.tabAvailabilitiesOverview.Padding = new System.Windows.Forms.Padding(6);
			this.tabAvailabilitiesOverview.Size = new System.Drawing.Size(1692, 1240);
			this.tabAvailabilitiesOverview.TabIndex = 0;
			this.tabAvailabilitiesOverview.Text = "Availabilities Overview";
			this.tabAvailabilitiesOverview.UseVisualStyleBackColor = true;
			// 
			// excelDataGrid
			// 
			this.excelDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.excelDataGrid.Location = new System.Drawing.Point(0, 0);
			this.excelDataGrid.Margin = new System.Windows.Forms.Padding(6);
			this.excelDataGrid.Name = "excelDataGrid";
			this.excelDataGrid.RowHeadersWidth = 82;
			this.excelDataGrid.Size = new System.Drawing.Size(1692, 1237);
			this.excelDataGrid.TabIndex = 0;
			// 
			// tabAvailabilitiesMap
			// 
			this.tabAvailabilitiesMap.Controls.Add(this.lblPropertyList);
			this.tabAvailabilitiesMap.Controls.Add(this.txtFileList);
			this.tabAvailabilitiesMap.Controls.Add(this.rdoGeneralImage);
			this.tabAvailabilitiesMap.Controls.Add(this.rdoFloorPlan);
			this.tabAvailabilitiesMap.Controls.Add(this.btnDeleteImage);
			this.tabAvailabilitiesMap.Controls.Add(this.btnSaveImage);
			this.tabAvailabilitiesMap.Controls.Add(this.btnSelectImage);
			this.tabAvailabilitiesMap.Controls.Add(this.propertyImage);
			this.tabAvailabilitiesMap.Controls.Add(this.txtPropertyImages);
			this.tabAvailabilitiesMap.Controls.Add(this.txtFloorPlansExist);
			this.tabAvailabilitiesMap.Controls.Add(this.lblPropertyImages);
			this.tabAvailabilitiesMap.Controls.Add(this.lblFloorPlans);
			this.tabAvailabilitiesMap.Controls.Add(this.cboProperties);
			this.tabAvailabilitiesMap.Location = new System.Drawing.Point(8, 39);
			this.tabAvailabilitiesMap.Margin = new System.Windows.Forms.Padding(6);
			this.tabAvailabilitiesMap.Name = "tabAvailabilitiesMap";
			this.tabAvailabilitiesMap.Padding = new System.Windows.Forms.Padding(6);
			this.tabAvailabilitiesMap.Size = new System.Drawing.Size(1692, 1240);
			this.tabAvailabilitiesMap.TabIndex = 1;
			this.tabAvailabilitiesMap.Text = "Properties";
			this.tabAvailabilitiesMap.UseVisualStyleBackColor = true;
			// 
			// lblPropertyList
			// 
			this.lblPropertyList.AutoSize = true;
			this.lblPropertyList.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblPropertyList.Location = new System.Drawing.Point(1024, 202);
			this.lblPropertyList.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblPropertyList.Name = "lblPropertyList";
			this.lblPropertyList.Size = new System.Drawing.Size(319, 36);
			this.lblPropertyList.TabIndex = 11;
			this.lblPropertyList.Text = "List of Property Images";
			// 
			// txtFileList
			// 
			this.txtFileList.Location = new System.Drawing.Point(1020, 244);
			this.txtFileList.Margin = new System.Windows.Forms.Padding(6);
			this.txtFileList.Multiline = true;
			this.txtFileList.Name = "txtFileList";
			this.txtFileList.Size = new System.Drawing.Size(660, 512);
			this.txtFileList.TabIndex = 10;
			// 
			// rdoGeneralImage
			// 
			this.rdoGeneralImage.AutoSize = true;
			this.rdoGeneralImage.Location = new System.Drawing.Point(774, 306);
			this.rdoGeneralImage.Margin = new System.Windows.Forms.Padding(6);
			this.rdoGeneralImage.Name = "rdoGeneralImage";
			this.rdoGeneralImage.Size = new System.Drawing.Size(183, 29);
			this.rdoGeneralImage.TabIndex = 9;
			this.rdoGeneralImage.TabStop = true;
			this.rdoGeneralImage.Text = "General Image";
			this.rdoGeneralImage.UseVisualStyleBackColor = true;
			this.rdoGeneralImage.CheckedChanged += new System.EventHandler(this.rdoGeneralImage_CheckedChanged);
			// 
			// rdoFloorPlan
			// 
			this.rdoFloorPlan.AutoSize = true;
			this.rdoFloorPlan.Location = new System.Drawing.Point(774, 262);
			this.rdoFloorPlan.Margin = new System.Windows.Forms.Padding(6);
			this.rdoFloorPlan.Name = "rdoFloorPlan";
			this.rdoFloorPlan.Size = new System.Drawing.Size(141, 29);
			this.rdoFloorPlan.TabIndex = 8;
			this.rdoFloorPlan.TabStop = true;
			this.rdoFloorPlan.Text = "Floor Plan";
			this.rdoFloorPlan.UseVisualStyleBackColor = true;
			this.rdoFloorPlan.CheckedChanged += new System.EventHandler(this.rdoFloorPlan_CheckedChanged);
			// 
			// btnDeleteImage
			// 
			this.btnDeleteImage.Location = new System.Drawing.Point(574, 188);
			this.btnDeleteImage.Margin = new System.Windows.Forms.Padding(6);
			this.btnDeleteImage.Name = "btnDeleteImage";
			this.btnDeleteImage.Size = new System.Drawing.Size(188, 44);
			this.btnDeleteImage.TabIndex = 7;
			this.btnDeleteImage.Text = "Delete";
			this.btnDeleteImage.UseVisualStyleBackColor = true;
			this.btnDeleteImage.Click += new System.EventHandler(this.btnDeleteImage_Click);
			// 
			// btnSaveImage
			// 
			this.btnSaveImage.Location = new System.Drawing.Point(288, 188);
			this.btnSaveImage.Margin = new System.Windows.Forms.Padding(6);
			this.btnSaveImage.Name = "btnSaveImage";
			this.btnSaveImage.Size = new System.Drawing.Size(274, 44);
			this.btnSaveImage.TabIndex = 6;
			this.btnSaveImage.Text = "Save";
			this.btnSaveImage.UseVisualStyleBackColor = true;
			this.btnSaveImage.Click += new System.EventHandler(this.btnSaveImage_Click);
			// 
			// btnSelectImage
			// 
			this.btnSelectImage.Location = new System.Drawing.Point(17, 188);
			this.btnSelectImage.Margin = new System.Windows.Forms.Padding(6);
			this.btnSelectImage.Name = "btnSelectImage";
			this.btnSelectImage.Size = new System.Drawing.Size(259, 44);
			this.btnSelectImage.TabIndex = 5;
			this.btnSelectImage.Text = "Select ";
			this.btnSelectImage.UseVisualStyleBackColor = true;
			this.btnSelectImage.Click += new System.EventHandler(this.btnSelectImage_Click);
			// 
			// propertyImage
			// 
			this.propertyImage.Image = ((System.Drawing.Image)(resources.GetObject("propertyImage.Image")));
			this.propertyImage.InitialImage = ((System.Drawing.Image)(resources.GetObject("propertyImage.InitialImage")));
			this.propertyImage.Location = new System.Drawing.Point(12, 244);
			this.propertyImage.Margin = new System.Windows.Forms.Padding(6);
			this.propertyImage.Name = "propertyImage";
			this.propertyImage.Size = new System.Drawing.Size(750, 515);
			this.propertyImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
			this.propertyImage.TabIndex = 3;
			this.propertyImage.TabStop = false;
			// 
			// txtPropertyImages
			// 
			this.txtPropertyImages.Enabled = false;
			this.txtPropertyImages.Location = new System.Drawing.Point(282, 120);
			this.txtPropertyImages.Margin = new System.Windows.Forms.Padding(6);
			this.txtPropertyImages.Name = "txtPropertyImages";
			this.txtPropertyImages.Size = new System.Drawing.Size(62, 31);
			this.txtPropertyImages.TabIndex = 2;
			// 
			// txtFloorPlansExist
			// 
			this.txtFloorPlansExist.Enabled = false;
			this.txtFloorPlansExist.Location = new System.Drawing.Point(282, 69);
			this.txtFloorPlansExist.Margin = new System.Windows.Forms.Padding(6);
			this.txtFloorPlansExist.Name = "txtFloorPlansExist";
			this.txtFloorPlansExist.Size = new System.Drawing.Size(62, 31);
			this.txtFloorPlansExist.TabIndex = 2;
			// 
			// lblPropertyImages
			// 
			this.lblPropertyImages.AutoSize = true;
			this.lblPropertyImages.Location = new System.Drawing.Point(4, 123);
			this.lblPropertyImages.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblPropertyImages.Name = "lblPropertyImages";
			this.lblPropertyImages.Size = new System.Drawing.Size(168, 25);
			this.lblPropertyImages.TabIndex = 1;
			this.lblPropertyImages.Text = "Property Images";
			// 
			// lblFloorPlans
			// 
			this.lblFloorPlans.AutoSize = true;
			this.lblFloorPlans.Location = new System.Drawing.Point(12, 72);
			this.lblFloorPlans.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblFloorPlans.Name = "lblFloorPlans";
			this.lblFloorPlans.Size = new System.Drawing.Size(121, 25);
			this.lblFloorPlans.TabIndex = 1;
			this.lblFloorPlans.Text = "Floor Plans";
			// 
			// cboProperties
			// 
			this.cboProperties.FormattingEnabled = true;
			this.cboProperties.Location = new System.Drawing.Point(12, 12);
			this.cboProperties.Margin = new System.Windows.Forms.Padding(6);
			this.cboProperties.Name = "cboProperties";
			this.cboProperties.Size = new System.Drawing.Size(680, 33);
			this.cboProperties.TabIndex = 0;
			this.cboProperties.Text = "-Select Property to Add Images-";
			this.cboProperties.TextChanged += new System.EventHandler(this.cboProperties_TextChanged);
			// 
			// tabNeeds
			// 
			this.tabNeeds.Controls.Add(this.btnSaveNeeds);
			this.tabNeeds.Controls.Add(this.txtClient);
			this.tabNeeds.Controls.Add(this.txtParking);
			this.tabNeeds.Controls.Add(this.txtComments);
			this.tabNeeds.Controls.Add(this.txtBuildingType);
			this.tabNeeds.Controls.Add(this.txtLocation);
			this.tabNeeds.Controls.Add(this.txtRelocationObjectives);
			this.tabNeeds.Controls.Add(this.txtGrowth);
			this.tabNeeds.Controls.Add(this.txtTerm);
			this.tabNeeds.Controls.Add(this.txtArea);
			this.tabNeeds.Controls.Add(this.lblClientName);
			this.tabNeeds.Controls.Add(this.lblComments);
			this.tabNeeds.Controls.Add(this.lblParking);
			this.tabNeeds.Controls.Add(this.lblBuildingType);
			this.tabNeeds.Controls.Add(this.lblLocation);
			this.tabNeeds.Controls.Add(this.lblRelocateObjectives);
			this.tabNeeds.Controls.Add(this.lblGrowth);
			this.tabNeeds.Controls.Add(this.lblTerm);
			this.tabNeeds.Controls.Add(this.lblArea);
			this.tabNeeds.Location = new System.Drawing.Point(8, 39);
			this.tabNeeds.Margin = new System.Windows.Forms.Padding(6);
			this.tabNeeds.Name = "tabNeeds";
			this.tabNeeds.Size = new System.Drawing.Size(1692, 1240);
			this.tabNeeds.TabIndex = 2;
			this.tabNeeds.Text = "Needs Analysis";
			this.tabNeeds.UseVisualStyleBackColor = true;
			// 
			// btnSaveNeeds
			// 
			this.btnSaveNeeds.Enabled = false;
			this.btnSaveNeeds.Location = new System.Drawing.Point(1098, 683);
			this.btnSaveNeeds.Margin = new System.Windows.Forms.Padding(6);
			this.btnSaveNeeds.Name = "btnSaveNeeds";
			this.btnSaveNeeds.Size = new System.Drawing.Size(150, 44);
			this.btnSaveNeeds.TabIndex = 23;
			this.btnSaveNeeds.Text = "Save Needs";
			this.btnSaveNeeds.UseVisualStyleBackColor = true;
			this.btnSaveNeeds.Click += new System.EventHandler(this.btnSaveNeeds_Click);
			// 
			// txtClient
			// 
			this.txtClient.Location = new System.Drawing.Point(720, 690);
			this.txtClient.Margin = new System.Windows.Forms.Padding(6);
			this.txtClient.Name = "txtClient";
			this.txtClient.Size = new System.Drawing.Size(328, 31);
			this.txtClient.TabIndex = 6;
			// 
			// txtParking
			// 
			this.txtParking.Location = new System.Drawing.Point(720, 361);
			this.txtParking.Margin = new System.Windows.Forms.Padding(6);
			this.txtParking.Name = "txtParking";
			this.txtParking.Size = new System.Drawing.Size(328, 31);
			this.txtParking.TabIndex = 5;
			// 
			// txtComments
			// 
			this.txtComments.Location = new System.Drawing.Point(720, 421);
			this.txtComments.Margin = new System.Windows.Forms.Padding(6);
			this.txtComments.Multiline = true;
			this.txtComments.Name = "txtComments";
			this.txtComments.Size = new System.Drawing.Size(528, 235);
			this.txtComments.TabIndex = 4;
			// 
			// txtBuildingType
			// 
			this.txtBuildingType.Location = new System.Drawing.Point(720, 311);
			this.txtBuildingType.Margin = new System.Windows.Forms.Padding(6);
			this.txtBuildingType.Name = "txtBuildingType";
			this.txtBuildingType.Size = new System.Drawing.Size(328, 31);
			this.txtBuildingType.TabIndex = 3;
			// 
			// txtLocation
			// 
			this.txtLocation.Location = new System.Drawing.Point(720, 255);
			this.txtLocation.Margin = new System.Windows.Forms.Padding(6);
			this.txtLocation.Name = "txtLocation";
			this.txtLocation.Size = new System.Drawing.Size(328, 31);
			this.txtLocation.TabIndex = 3;
			// 
			// txtRelocationObjectives
			// 
			this.txtRelocationObjectives.Location = new System.Drawing.Point(720, 200);
			this.txtRelocationObjectives.Margin = new System.Windows.Forms.Padding(6);
			this.txtRelocationObjectives.Name = "txtRelocationObjectives";
			this.txtRelocationObjectives.Size = new System.Drawing.Size(328, 31);
			this.txtRelocationObjectives.TabIndex = 2;
			// 
			// txtGrowth
			// 
			this.txtGrowth.Location = new System.Drawing.Point(720, 144);
			this.txtGrowth.Margin = new System.Windows.Forms.Padding(6);
			this.txtGrowth.Name = "txtGrowth";
			this.txtGrowth.Size = new System.Drawing.Size(328, 31);
			this.txtGrowth.TabIndex = 1;
			// 
			// txtTerm
			// 
			this.txtTerm.Location = new System.Drawing.Point(720, 88);
			this.txtTerm.Margin = new System.Windows.Forms.Padding(6);
			this.txtTerm.Name = "txtTerm";
			this.txtTerm.Size = new System.Drawing.Size(328, 31);
			this.txtTerm.TabIndex = 1;
			// 
			// txtArea
			// 
			this.txtArea.Location = new System.Drawing.Point(720, 40);
			this.txtArea.Margin = new System.Windows.Forms.Padding(6);
			this.txtArea.Name = "txtArea";
			this.txtArea.Size = new System.Drawing.Size(328, 31);
			this.txtArea.TabIndex = 1;
			// 
			// lblClientName
			// 
			this.lblClientName.AutoSize = true;
			this.lblClientName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblClientName.Location = new System.Drawing.Point(434, 690);
			this.lblClientName.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblClientName.Name = "lblClientName";
			this.lblClientName.Size = new System.Drawing.Size(84, 31);
			this.lblClientName.TabIndex = 0;
			this.lblClientName.Text = "Client";
			// 
			// lblComments
			// 
			this.lblComments.AutoSize = true;
			this.lblComments.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblComments.Location = new System.Drawing.Point(434, 421);
			this.lblComments.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblComments.Name = "lblComments";
			this.lblComments.Size = new System.Drawing.Size(145, 31);
			this.lblComments.TabIndex = 0;
			this.lblComments.Text = "Comments";
			// 
			// lblParking
			// 
			this.lblParking.AutoSize = true;
			this.lblParking.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblParking.Location = new System.Drawing.Point(434, 361);
			this.lblParking.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblParking.Name = "lblParking";
			this.lblParking.Size = new System.Drawing.Size(106, 31);
			this.lblParking.TabIndex = 0;
			this.lblParking.Text = "Parking";
			// 
			// lblBuildingType
			// 
			this.lblBuildingType.AutoSize = true;
			this.lblBuildingType.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblBuildingType.Location = new System.Drawing.Point(434, 308);
			this.lblBuildingType.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblBuildingType.Name = "lblBuildingType";
			this.lblBuildingType.Size = new System.Drawing.Size(178, 31);
			this.lblBuildingType.TabIndex = 0;
			this.lblBuildingType.Text = "Building Type";
			// 
			// lblLocation
			// 
			this.lblLocation.AutoSize = true;
			this.lblLocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblLocation.Location = new System.Drawing.Point(434, 255);
			this.lblLocation.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblLocation.Name = "lblLocation";
			this.lblLocation.Size = new System.Drawing.Size(117, 31);
			this.lblLocation.TabIndex = 0;
			this.lblLocation.Text = "Location";
			// 
			// lblRelocateObjectives
			// 
			this.lblRelocateObjectives.AutoSize = true;
			this.lblRelocateObjectives.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblRelocateObjectives.Location = new System.Drawing.Point(434, 205);
			this.lblRelocateObjectives.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblRelocateObjectives.Name = "lblRelocateObjectives";
			this.lblRelocateObjectives.Size = new System.Drawing.Size(278, 31);
			this.lblRelocateObjectives.TabIndex = 0;
			this.lblRelocateObjectives.Text = "Relocation Objectives";
			// 
			// lblGrowth
			// 
			this.lblGrowth.AutoSize = true;
			this.lblGrowth.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblGrowth.Location = new System.Drawing.Point(434, 144);
			this.lblGrowth.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblGrowth.Name = "lblGrowth";
			this.lblGrowth.Size = new System.Drawing.Size(102, 31);
			this.lblGrowth.TabIndex = 0;
			this.lblGrowth.Text = "Growth";
			// 
			// lblTerm
			// 
			this.lblTerm.AutoSize = true;
			this.lblTerm.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTerm.Location = new System.Drawing.Point(434, 88);
			this.lblTerm.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblTerm.Name = "lblTerm";
			this.lblTerm.Size = new System.Drawing.Size(77, 31);
			this.lblTerm.TabIndex = 0;
			this.lblTerm.Text = "Term";
			// 
			// lblArea
			// 
			this.lblArea.AutoSize = true;
			this.lblArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblArea.Location = new System.Drawing.Point(434, 40);
			this.lblArea.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lblArea.Name = "lblArea";
			this.lblArea.Size = new System.Drawing.Size(71, 31);
			this.lblArea.TabIndex = 0;
			this.lblArea.Text = "Area";
			// 
			// openFileDialog
			// 
			this.openFileDialog.FileName = "openFileDialog";
			// 
			// btnGenerateSurvey
			// 
			this.btnGenerateSurvey.Enabled = false;
			this.btnGenerateSurvey.Location = new System.Drawing.Point(484, 48);
			this.btnGenerateSurvey.Margin = new System.Windows.Forms.Padding(6);
			this.btnGenerateSurvey.Name = "btnGenerateSurvey";
			this.btnGenerateSurvey.Size = new System.Drawing.Size(256, 71);
			this.btnGenerateSurvey.TabIndex = 3;
			this.btnGenerateSurvey.Text = "Create Survey";
			this.btnGenerateSurvey.UseVisualStyleBackColor = true;
			this.btnGenerateSurvey.Click += new System.EventHandler(this.btnGenerateSurvey_Click);
			// 
			// cboLanguage
			// 
			this.cboLanguage.FormattingEnabled = true;
			this.cboLanguage.Items.AddRange(new object[] {
            "English",
            "French"});
			this.cboLanguage.Location = new System.Drawing.Point(786, 48);
			this.cboLanguage.Margin = new System.Windows.Forms.Padding(6);
			this.cboLanguage.Name = "cboLanguage";
			this.cboLanguage.Size = new System.Drawing.Size(486, 33);
			this.cboLanguage.TabIndex = 4;
			this.cboLanguage.Text = "-Select a Language/sélectionnez une langue-";
			this.cboLanguage.SelectedIndexChanged += new System.EventHandler(this.cboLanguage_SelectedIndexChanged);
			// 
			// chkRem
			// 
			this.chkRem.AutoSize = true;
			this.chkRem.Location = new System.Drawing.Point(786, 90);
			this.chkRem.Name = "chkRem";
			this.chkRem.Size = new System.Drawing.Size(103, 29);
			this.chkRem.TabIndex = 5;
			this.chkRem.Text = "REM?";
			this.chkRem.UseVisualStyleBackColor = true;
			// 
			// LoadExcelSurvey
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoSize = true;
			this.ClientSize = new System.Drawing.Size(2276, 1487);
			this.Controls.Add(this.chkRem);
			this.Controls.Add(this.cboLanguage);
			this.Controls.Add(this.btnGenerateSurvey);
			this.Controls.Add(this.documentPreview);
			this.Controls.Add(this.btnLoadSurvey);
			this.Margin = new System.Windows.Forms.Padding(6);
			this.Name = "LoadExcelSurvey";
			this.Text = "Terramont Survey Generator";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LoadExcelSurvey_FormClosing);
			this.Load += new System.EventHandler(this.LoadExcelSurvey_Load);
			this.documentPreview.ResumeLayout(false);
			this.tabAvailabilitiesOverview.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.excelDataGrid)).EndInit();
			this.tabAvailabilitiesMap.ResumeLayout(false);
			this.tabAvailabilitiesMap.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.propertyImage)).EndInit();
			this.tabNeeds.ResumeLayout(false);
			this.tabNeeds.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnLoadSurvey;
		private System.Windows.Forms.TabControl documentPreview;
		private System.Windows.Forms.TabPage tabAvailabilitiesOverview;
		private System.Windows.Forms.DataGridView excelDataGrid;
		private System.Windows.Forms.TabPage tabAvailabilitiesMap;
		private System.Windows.Forms.ComboBox cboProperties;
		private System.Windows.Forms.Label lblFloorPlans;
		private System.Windows.Forms.TextBox txtFloorPlansExist;
		private System.Windows.Forms.PictureBox propertyImage;
		private System.Windows.Forms.Button btnSelectImage;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Button btnDeleteImage;
		private System.Windows.Forms.Button btnSaveImage;
		private System.Windows.Forms.RadioButton rdoGeneralImage;
		private System.Windows.Forms.RadioButton rdoFloorPlan;
		private System.Windows.Forms.TextBox txtPropertyImages;
		private System.Windows.Forms.Label lblPropertyImages;
		private System.Windows.Forms.Label lblPropertyList;
		private System.Windows.Forms.TextBox txtFileList;
		private System.Windows.Forms.TabPage tabNeeds;
		private System.Windows.Forms.TextBox txtParking;
		private System.Windows.Forms.TextBox txtComments;
		private System.Windows.Forms.TextBox txtBuildingType;
		private System.Windows.Forms.TextBox txtLocation;
		private System.Windows.Forms.TextBox txtRelocationObjectives;
		private System.Windows.Forms.TextBox txtGrowth;
		private System.Windows.Forms.TextBox txtTerm;
		private System.Windows.Forms.TextBox txtArea;
		private System.Windows.Forms.Label lblComments;
		private System.Windows.Forms.Label lblParking;
		private System.Windows.Forms.Label lblBuildingType;
		private System.Windows.Forms.Label lblLocation;
		private System.Windows.Forms.Label lblRelocateObjectives;
		private System.Windows.Forms.Label lblGrowth;
		private System.Windows.Forms.Label lblTerm;
		private System.Windows.Forms.Label lblArea;
		private System.Windows.Forms.TextBox txtClient;
		private System.Windows.Forms.Label lblClientName;
		private System.Windows.Forms.Button btnSaveNeeds;
		private System.Windows.Forms.Button btnGenerateSurvey;
		private System.ComponentModel.BackgroundWorker bkgProcessFile;
		private System.Windows.Forms.ComboBox cboLanguage;
		private System.Windows.Forms.CheckBox chkRem;
	}
}

