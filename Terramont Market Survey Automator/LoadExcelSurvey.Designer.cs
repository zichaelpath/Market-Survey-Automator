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
            this.txtNumOfProperties = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboProperties = new System.Windows.Forms.ComboBox();
            this.propertyImage = new System.Windows.Forms.PictureBox();
            this.btnFirstImage = new System.Windows.Forms.Button();
            this.btnPreviousImage = new System.Windows.Forms.Button();
            this.btnNextImage = new System.Windows.Forms.Button();
            this.btnLastImage = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnSelectImage = new System.Windows.Forms.Button();
            this.documentPreview.SuspendLayout();
            this.tabAvailabilitiesOverview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.excelDataGrid)).BeginInit();
            this.tabAvailabilitiesMap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.propertyImage)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLoadSurvey
            // 
            this.btnLoadSurvey.Location = new System.Drawing.Point(91, 25);
            this.btnLoadSurvey.Name = "btnLoadSurvey";
            this.btnLoadSurvey.Size = new System.Drawing.Size(120, 37);
            this.btnLoadSurvey.TabIndex = 1;
            this.btnLoadSurvey.Text = "Load Excel Survey";
            this.btnLoadSurvey.UseVisualStyleBackColor = true;
            this.btnLoadSurvey.Click += new System.EventHandler(this.btnLoadSurvey_Click);
            // 
            // documentPreview
            // 
            this.documentPreview.Controls.Add(this.tabAvailabilitiesOverview);
            this.documentPreview.Controls.Add(this.tabAvailabilitiesMap);
            this.documentPreview.Location = new System.Drawing.Point(35, 92);
            this.documentPreview.Name = "documentPreview";
            this.documentPreview.SelectedIndex = 0;
            this.documentPreview.Size = new System.Drawing.Size(777, 406);
            this.documentPreview.TabIndex = 2;
            // 
            // tabAvailabilitiesOverview
            // 
            this.tabAvailabilitiesOverview.Controls.Add(this.excelDataGrid);
            this.tabAvailabilitiesOverview.Location = new System.Drawing.Point(4, 22);
            this.tabAvailabilitiesOverview.Name = "tabAvailabilitiesOverview";
            this.tabAvailabilitiesOverview.Padding = new System.Windows.Forms.Padding(3);
            this.tabAvailabilitiesOverview.Size = new System.Drawing.Size(769, 380);
            this.tabAvailabilitiesOverview.TabIndex = 0;
            this.tabAvailabilitiesOverview.Text = "Availabilities Overview";
            this.tabAvailabilitiesOverview.UseVisualStyleBackColor = true;
            // 
            // excelDataGrid
            // 
            this.excelDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.excelDataGrid.Location = new System.Drawing.Point(0, 0);
            this.excelDataGrid.Name = "excelDataGrid";
            this.excelDataGrid.Size = new System.Drawing.Size(769, 380);
            this.excelDataGrid.TabIndex = 0;
            // 
            // tabAvailabilitiesMap
            // 
            this.tabAvailabilitiesMap.Controls.Add(this.btnSelectImage);
            this.tabAvailabilitiesMap.Controls.Add(this.btnLastImage);
            this.tabAvailabilitiesMap.Controls.Add(this.btnNextImage);
            this.tabAvailabilitiesMap.Controls.Add(this.btnPreviousImage);
            this.tabAvailabilitiesMap.Controls.Add(this.btnFirstImage);
            this.tabAvailabilitiesMap.Controls.Add(this.propertyImage);
            this.tabAvailabilitiesMap.Controls.Add(this.txtNumOfProperties);
            this.tabAvailabilitiesMap.Controls.Add(this.label1);
            this.tabAvailabilitiesMap.Controls.Add(this.cboProperties);
            this.tabAvailabilitiesMap.Location = new System.Drawing.Point(4, 22);
            this.tabAvailabilitiesMap.Name = "tabAvailabilitiesMap";
            this.tabAvailabilitiesMap.Padding = new System.Windows.Forms.Padding(3);
            this.tabAvailabilitiesMap.Size = new System.Drawing.Size(769, 380);
            this.tabAvailabilitiesMap.TabIndex = 1;
            this.tabAvailabilitiesMap.Text = "Availabilities Map";
            this.tabAvailabilitiesMap.UseVisualStyleBackColor = true;
            // 
            // txtNumOfProperties
            // 
            this.txtNumOfProperties.Enabled = false;
            this.txtNumOfProperties.Location = new System.Drawing.Point(185, 31);
            this.txtNumOfProperties.Name = "txtNumOfProperties";
            this.txtNumOfProperties.Size = new System.Drawing.Size(33, 20);
            this.txtNumOfProperties.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(176, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Total Images For Selected Property:";
            // 
            // cboProperties
            // 
            this.cboProperties.FormattingEnabled = true;
            this.cboProperties.Location = new System.Drawing.Point(6, 6);
            this.cboProperties.Name = "cboProperties";
            this.cboProperties.Size = new System.Drawing.Size(342, 21);
            this.cboProperties.TabIndex = 0;
            this.cboProperties.Text = "-Select Property to Add Images-";
            // 
            // propertyImage
            // 
            this.propertyImage.Image = ((System.Drawing.Image)(resources.GetObject("propertyImage.Image")));
            this.propertyImage.InitialImage = ((System.Drawing.Image)(resources.GetObject("propertyImage.InitialImage")));
            this.propertyImage.Location = new System.Drawing.Point(9, 60);
            this.propertyImage.Name = "propertyImage";
            this.propertyImage.Size = new System.Drawing.Size(375, 268);
            this.propertyImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.propertyImage.TabIndex = 3;
            this.propertyImage.TabStop = false;
            // 
            // btnFirstImage
            // 
            this.btnFirstImage.Location = new System.Drawing.Point(26, 334);
            this.btnFirstImage.Name = "btnFirstImage";
            this.btnFirstImage.Size = new System.Drawing.Size(75, 23);
            this.btnFirstImage.TabIndex = 4;
            this.btnFirstImage.Text = "First Image";
            this.btnFirstImage.UseVisualStyleBackColor = true;
            // 
            // btnPreviousImage
            // 
            this.btnPreviousImage.Location = new System.Drawing.Point(107, 334);
            this.btnPreviousImage.Name = "btnPreviousImage";
            this.btnPreviousImage.Size = new System.Drawing.Size(92, 23);
            this.btnPreviousImage.TabIndex = 4;
            this.btnPreviousImage.Text = "Previous Image";
            this.btnPreviousImage.UseVisualStyleBackColor = true;
            // 
            // btnNextImage
            // 
            this.btnNextImage.Location = new System.Drawing.Point(205, 334);
            this.btnNextImage.Name = "btnNextImage";
            this.btnNextImage.Size = new System.Drawing.Size(76, 23);
            this.btnNextImage.TabIndex = 4;
            this.btnNextImage.Text = "Next Image";
            this.btnNextImage.UseVisualStyleBackColor = true;
            // 
            // btnLastImage
            // 
            this.btnLastImage.Location = new System.Drawing.Point(289, 334);
            this.btnLastImage.Name = "btnLastImage";
            this.btnLastImage.Size = new System.Drawing.Size(76, 23);
            this.btnLastImage.TabIndex = 4;
            this.btnLastImage.Text = "Last Image";
            this.btnLastImage.UseVisualStyleBackColor = true;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // btnSelectImage
            // 
            this.btnSelectImage.Location = new System.Drawing.Point(233, 30);
            this.btnSelectImage.Name = "btnSelectImage";
            this.btnSelectImage.Size = new System.Drawing.Size(82, 23);
            this.btnSelectImage.TabIndex = 5;
            this.btnSelectImage.Text = "Select Image";
            this.btnSelectImage.UseVisualStyleBackColor = true;
            this.btnSelectImage.Click += new System.EventHandler(this.btnSelectImage_Click);
            // 
            // LoadExcelSurvey
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(886, 510);
            this.Controls.Add(this.documentPreview);
            this.Controls.Add(this.btnLoadSurvey);
            this.Name = "LoadExcelSurvey";
            this.Text = "Terramont Survey Generator";
            this.Load += new System.EventHandler(this.LoadExcelSurvey_Load);
            this.documentPreview.ResumeLayout(false);
            this.tabAvailabilitiesOverview.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.excelDataGrid)).EndInit();
            this.tabAvailabilitiesMap.ResumeLayout(false);
            this.tabAvailabilitiesMap.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.propertyImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnLoadSurvey;
        private System.Windows.Forms.TabControl documentPreview;
        private System.Windows.Forms.TabPage tabAvailabilitiesOverview;
        private System.Windows.Forms.DataGridView excelDataGrid;
        private System.Windows.Forms.TabPage tabAvailabilitiesMap;
        private System.Windows.Forms.ComboBox cboProperties;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtNumOfProperties;
        private System.Windows.Forms.PictureBox propertyImage;
        private System.Windows.Forms.Button btnLastImage;
        private System.Windows.Forms.Button btnNextImage;
        private System.Windows.Forms.Button btnPreviousImage;
        private System.Windows.Forms.Button btnFirstImage;
        private System.Windows.Forms.Button btnSelectImage;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
    }
}

