
namespace MartinAppGUI
{
    partial class DataHandlerApp
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DataHandlerApp));
            this.openFile = new System.Windows.Forms.OpenFileDialog();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panel1 = new System.Windows.Forms.Panel();
            this.bttnStop = new System.Windows.Forms.Button();
            this.bttnOpen = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.bttnClose = new System.Windows.Forms.Button();
            this.bttnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFile
            // 
            this.openFile.FileName = "openFile";
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.Snow;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 137);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(976, 522);
            this.dataGridView1.TabIndex = 2;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // backgroundWorker2
            // 
            this.backgroundWorker2.WorkerReportsProgress = true;
            this.backgroundWorker2.WorkerSupportsCancellation = true;
            this.backgroundWorker2.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker2_DoWork);
            this.backgroundWorker2.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker2_ProgressChanged);
            this.backgroundWorker2.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker2_RunWorkerCompleted);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(13, 665);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(975, 23);
            this.progressBar1.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.bttnStop);
            this.panel1.Controls.Add(this.bttnOpen);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.bttnClose);
            this.panel1.Controls.Add(this.bttnExport);
            this.panel1.Location = new System.Drawing.Point(-1, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1001, 131);
            this.panel1.TabIndex = 4;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            // 
            // bttnStop
            // 
            this.bttnStop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(50)))), ((int)(((byte)(90)))));
            this.bttnStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bttnStop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttnStop.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.bttnStop.ForeColor = System.Drawing.Color.White;
            this.bttnStop.Location = new System.Drawing.Point(765, 61);
            this.bttnStop.Name = "bttnStop";
            this.bttnStop.Size = new System.Drawing.Size(192, 54);
            this.bttnStop.TabIndex = 0;
            this.bttnStop.Text = "✋ Leállít";
            this.bttnStop.UseVisualStyleBackColor = false;
            this.bttnStop.Click += new System.EventHandler(this.button1_Click);
            this.bttnStop.MouseEnter += new System.EventHandler(this.bttnStop_MouseEnter);
            this.bttnStop.MouseLeave += new System.EventHandler(this.bttnStop_MouseLeave);
            // 
            // bttnOpen
            // 
            this.bttnOpen.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(50)))), ((int)(((byte)(90)))));
            this.bttnOpen.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bttnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttnOpen.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.bttnOpen.ForeColor = System.Drawing.Color.White;
            this.bttnOpen.Location = new System.Drawing.Point(369, 61);
            this.bttnOpen.Name = "bttnOpen";
            this.bttnOpen.Size = new System.Drawing.Size(192, 54);
            this.bttnOpen.TabIndex = 0;
            this.bttnOpen.Text = "📁 XLSX betöltése";
            this.bttnOpen.UseVisualStyleBackColor = false;
            this.bttnOpen.Click += new System.EventHandler(this.megnyitas_Click);
            this.bttnOpen.MouseEnter += new System.EventHandler(this.bttnOpen_MouseEnter_1);
            this.bttnOpen.MouseLeave += new System.EventHandler(this.bttnOpen_MouseLeave);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(50)))), ((int)(((byte)(90)))));
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(916, 7);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(35, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "_";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.button1.MouseEnter += new System.EventHandler(this.button1_MouseEnter);
            this.button1.MouseLeave += new System.EventHandler(this.button1_MouseLeave);
            // 
            // bttnClose
            // 
            this.bttnClose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(50)))), ((int)(((byte)(90)))));
            this.bttnClose.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bttnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttnClose.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.bttnClose.ForeColor = System.Drawing.Color.White;
            this.bttnClose.Location = new System.Drawing.Point(957, 7);
            this.bttnClose.Name = "bttnClose";
            this.bttnClose.Size = new System.Drawing.Size(35, 28);
            this.bttnClose.TabIndex = 0;
            this.bttnClose.Text = "X";
            this.bttnClose.UseVisualStyleBackColor = false;
            this.bttnClose.Click += new System.EventHandler(this.bttnClose_Click);
            this.bttnClose.MouseEnter += new System.EventHandler(this.bttnClose_MouseEnter);
            this.bttnClose.MouseLeave += new System.EventHandler(this.bttnClose_MouseLeave);
            // 
            // bttnExport
            // 
            this.bttnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(50)))), ((int)(((byte)(90)))));
            this.bttnExport.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bttnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttnExport.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.bttnExport.ForeColor = System.Drawing.Color.White;
            this.bttnExport.Location = new System.Drawing.Point(567, 61);
            this.bttnExport.Name = "bttnExport";
            this.bttnExport.Size = new System.Drawing.Size(192, 54);
            this.bttnExport.TabIndex = 0;
            this.bttnExport.Text = "💾 Adatok exportálása";
            this.bttnExport.UseVisualStyleBackColor = false;
            this.bttnExport.Click += new System.EventHandler(this.exportbutton_Click);
            this.bttnExport.MouseEnter += new System.EventHandler(this.bttnExport_MouseEnter);
            this.bttnExport.MouseLeave += new System.EventHandler(this.bttnExport_MouseLeave);
            // 
            // DataHandlerApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(1000, 700);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dataGridView1);
            this.DoubleBuffered = true;
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1000, 700);
            this.Name = "DataHandlerApp";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Adatkezelő App";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFile;
        private System.Windows.Forms.Button bttnOpen;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button bttnExport;
        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
        private System.Windows.Forms.Button bttnStop;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button bttnClose;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
    }
}

