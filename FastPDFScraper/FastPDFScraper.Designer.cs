namespace FastPDFScraper
{
    partial class FastPDFScraper
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
            this.btnOpenPDFFolder = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.chkMultiResult = new System.Windows.Forms.CheckBox();
            this.btnOpenKeywordFile = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.btnPrintResult = new System.Windows.Forms.Button();
            this.progressReport = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOpenPDFFolder
            // 
            this.btnOpenPDFFolder.Location = new System.Drawing.Point(12, 37);
            this.btnOpenPDFFolder.Name = "btnOpenPDFFolder";
            this.btnOpenPDFFolder.Size = new System.Drawing.Size(140, 74);
            this.btnOpenPDFFolder.TabIndex = 0;
            this.btnOpenPDFFolder.Text = "OpenPDFFolder";
            this.btnOpenPDFFolder.UseVisualStyleBackColor = true;
            this.btnOpenPDFFolder.Click += new System.EventHandler(this.btnOpenPDFFolder_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(439, 37);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(140, 74);
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // chkMultiResult
            // 
            this.chkMultiResult.AutoSize = true;
            this.chkMultiResult.Checked = true;
            this.chkMultiResult.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMultiResult.Location = new System.Drawing.Point(255, 249);
            this.chkMultiResult.Name = "chkMultiResult";
            this.chkMultiResult.Size = new System.Drawing.Size(81, 17);
            this.chkMultiResult.TabIndex = 2;
            this.chkMultiResult.Text = "Multi-Result";
            this.chkMultiResult.UseVisualStyleBackColor = true;
            // 
            // btnOpenKeywordFile
            // 
            this.btnOpenKeywordFile.Location = new System.Drawing.Point(229, 37);
            this.btnOpenKeywordFile.Name = "btnOpenKeywordFile";
            this.btnOpenKeywordFile.Size = new System.Drawing.Size(140, 74);
            this.btnOpenKeywordFile.TabIndex = 3;
            this.btnOpenKeywordFile.Text = "OpenKeywordFile";
            this.btnOpenKeywordFile.UseVisualStyleBackColor = true;
            this.btnOpenKeywordFile.Click += new System.EventHandler(this.btnOpenKeywordFile_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 367);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(567, 23);
            this.progressBar.TabIndex = 4;
            // 
            // btnPrintResult
            // 
            this.btnPrintResult.Location = new System.Drawing.Point(229, 134);
            this.btnPrintResult.Name = "btnPrintResult";
            this.btnPrintResult.Size = new System.Drawing.Size(140, 74);
            this.btnPrintResult.TabIndex = 5;
            this.btnPrintResult.Text = "PrintResult";
            this.btnPrintResult.UseVisualStyleBackColor = true;
            this.btnPrintResult.Click += new System.EventHandler(this.btnPrintResult_Click);
            // 
            // progressReport
            // 
            this.progressReport.AutoSize = true;
            this.progressReport.Location = new System.Drawing.Point(12, 326);
            this.progressReport.Name = "progressReport";
            this.progressReport.Size = new System.Drawing.Size(0, 13);
            this.progressReport.TabIndex = 6;
            // 
            // FastPDFScraper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 474);
            this.Controls.Add(this.progressReport);
            this.Controls.Add(this.btnPrintResult);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnOpenKeywordFile);
            this.Controls.Add(this.chkMultiResult);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnOpenPDFFolder);
            this.Name = "FastPDFScraper";
            this.Text = "FastPDFScraper";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenPDFFolder;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.CheckBox chkMultiResult;
        private System.Windows.Forms.Button btnOpenKeywordFile;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button btnPrintResult;
        private System.Windows.Forms.Label progressReport;
    }
}

