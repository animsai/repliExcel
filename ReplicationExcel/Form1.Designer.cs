﻿namespace ReplicationExcel
{
    partial class Form1
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.lsbNames = new System.Windows.Forms.ListBox();
            this.lblFileName = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btnBrowseNomsEleves = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cbxEncoding = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnDeleteName = new System.Windows.Forms.Button();
            this.btnSaveName = new System.Windows.Forms.Button();
            this.tbxAddName = new System.Windows.Forms.TextBox();
            this.btnAddName = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.tbxExcelFile = new System.Windows.Forms.TextBox();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lsbSheets = new System.Windows.Forms.ListBox();
            this.btnCopyExcel = new System.Windows.Forms.Button();
            this.lblFileExcel = new System.Windows.Forms.Label();
            this.opdFiles = new System.Windows.Forms.OpenFileDialog();
            this.sfdFile = new System.Windows.Forms.SaveFileDialog();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fichierToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lsbNames
            // 
            this.lsbNames.FormattingEnabled = true;
            this.lsbNames.Location = new System.Drawing.Point(6, 66);
            this.lsbNames.Name = "lsbNames";
            this.lsbNames.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lsbNames.Size = new System.Drawing.Size(419, 95);
            this.lsbNames.TabIndex = 0;
            this.lsbNames.SelectedIndexChanged += new System.EventHandler(this.lsbNames_SelectedIndexChanged);
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Location = new System.Drawing.Point(6, 26);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(0, 13);
            this.lblFileName.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.btnBrowseNomsEleves);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cbxEncoding);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnDeleteName);
            this.groupBox1.Controls.Add(this.btnSaveName);
            this.groupBox1.Controls.Add(this.tbxAddName);
            this.groupBox1.Controls.Add(this.btnAddName);
            this.groupBox1.Controls.Add(this.lsbNames);
            this.groupBox1.Controls.Add(this.lblFileName);
            this.groupBox1.Location = new System.Drawing.Point(0, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(434, 265);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Fichier contenant les noms d\'élève";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 176);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(84, 13);
            this.label9.TabIndex = 12;
            this.label9.Text = "Ajouter un nom :";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 26);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Classe existante :";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(101, 23);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(102, 21);
            this.comboBox1.TabIndex = 10;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // btnBrowseNomsEleves
            // 
            this.btnBrowseNomsEleves.Location = new System.Drawing.Point(320, 21);
            this.btnBrowseNomsEleves.Name = "btnBrowseNomsEleves";
            this.btnBrowseNomsEleves.Size = new System.Drawing.Size(105, 23);
            this.btnBrowseNomsEleves.TabIndex = 8;
            this.btnBrowseNomsEleves.Text = "ajouter des élèves";
            this.btnBrowseNomsEleves.UseVisualStyleBackColor = true;
            this.btnBrowseNomsEleves.Click += new System.EventHandler(this.btnBrowseNomsEleves_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(214, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Ajouter une classe :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(251, 233);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Codage du fichier :";
            // 
            // cbxEncoding
            // 
            this.cbxEncoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxEncoding.FormattingEnabled = true;
            this.cbxEncoding.Items.AddRange(new object[] {
            "UTF-8",
            "latin1"});
            this.cbxEncoding.Location = new System.Drawing.Point(353, 230);
            this.cbxEncoding.Name = "cbxEncoding";
            this.cbxEncoding.Size = new System.Drawing.Size(72, 21);
            this.cbxEncoding.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Noms :";
            // 
            // btnDeleteName
            // 
            this.btnDeleteName.Enabled = false;
            this.btnDeleteName.Location = new System.Drawing.Point(6, 199);
            this.btnDeleteName.Name = "btnDeleteName";
            this.btnDeleteName.Size = new System.Drawing.Size(419, 23);
            this.btnDeleteName.TabIndex = 5;
            this.btnDeleteName.Text = "Supprimer nom(s) séléctionné(s)";
            this.btnDeleteName.UseVisualStyleBackColor = true;
            this.btnDeleteName.Click += new System.EventHandler(this.btnDeleteName_Click);
            // 
            // btnSaveName
            // 
            this.btnSaveName.Enabled = false;
            this.btnSaveName.Location = new System.Drawing.Point(6, 228);
            this.btnSaveName.Name = "btnSaveName";
            this.btnSaveName.Size = new System.Drawing.Size(239, 23);
            this.btnSaveName.TabIndex = 4;
            this.btnSaveName.Text = "Sauvegarder les noms";
            this.btnSaveName.UseVisualStyleBackColor = true;
            this.btnSaveName.Click += new System.EventHandler(this.btnSaveName_Click);
            // 
            // tbxAddName
            // 
            this.tbxAddName.Location = new System.Drawing.Point(96, 173);
            this.tbxAddName.Name = "tbxAddName";
            this.tbxAddName.Size = new System.Drawing.Size(269, 20);
            this.tbxAddName.TabIndex = 3;
            this.tbxAddName.TextChanged += new System.EventHandler(this.tbxAddName_TextChanged);
            // 
            // btnAddName
            // 
            this.btnAddName.Enabled = false;
            this.btnAddName.Location = new System.Drawing.Point(371, 171);
            this.btnAddName.Name = "btnAddName";
            this.btnAddName.Size = new System.Drawing.Size(54, 23);
            this.btnAddName.TabIndex = 2;
            this.btnAddName.Text = "Ajouter";
            this.btnAddName.UseVisualStyleBackColor = true;
            this.btnAddName.Click += new System.EventHandler(this.btnAddName_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.listBox1);
            this.groupBox2.Controls.Add(this.tbxExcelFile);
            this.groupBox2.Controls.Add(this.btnBrowseExcel);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.lsbSheets);
            this.groupBox2.Controls.Add(this.btnCopyExcel);
            this.groupBox2.Controls.Add(this.lblFileExcel);
            this.groupBox2.Location = new System.Drawing.Point(440, 27);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(434, 265);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Fichier excel";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(7, 204);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(325, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "* veuilllez remplacer les champs des élèves par le mot : \"pokemon\" ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 128);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(113, 13);
            this.label7.TabIndex = 11;
            this.label7.Text = "Feuille recapitulative* :";
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(9, 144);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(419, 56);
            this.listBox1.TabIndex = 10;
            // 
            // tbxExcelFile
            // 
            this.tbxExcelFile.Location = new System.Drawing.Point(91, 23);
            this.tbxExcelFile.Name = "tbxExcelFile";
            this.tbxExcelFile.Size = new System.Drawing.Size(241, 20);
            this.tbxExcelFile.TabIndex = 9;
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Location = new System.Drawing.Point(338, 21);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(90, 23);
            this.btnBrowseExcel.TabIndex = 8;
            this.btnBrowseExcel.Text = "Parcourir...";
            this.btnBrowseExcel.UseVisualStyleBackColor = true;
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Feuille à cloner :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(79, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Fichier source :";
            // 
            // lsbSheets
            // 
            this.lsbSheets.FormattingEnabled = true;
            this.lsbSheets.Location = new System.Drawing.Point(9, 66);
            this.lsbSheets.Name = "lsbSheets";
            this.lsbSheets.Size = new System.Drawing.Size(419, 56);
            this.lsbSheets.TabIndex = 3;
            this.lsbSheets.SelectedIndexChanged += new System.EventHandler(this.lsbSheets_SelectedIndexChanged);
            // 
            // btnCopyExcel
            // 
            this.btnCopyExcel.Enabled = false;
            this.btnCopyExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCopyExcel.Location = new System.Drawing.Point(9, 224);
            this.btnCopyExcel.Name = "btnCopyExcel";
            this.btnCopyExcel.Size = new System.Drawing.Size(419, 30);
            this.btnCopyExcel.TabIndex = 2;
            this.btnCopyExcel.Text = "Générer feuilles excel";
            this.btnCopyExcel.UseVisualStyleBackColor = true;
            this.btnCopyExcel.Click += new System.EventHandler(this.btnCopyExcel_Click);
            // 
            // lblFileExcel
            // 
            this.lblFileExcel.AutoSize = true;
            this.lblFileExcel.Location = new System.Drawing.Point(6, 26);
            this.lblFileExcel.Name = "lblFileExcel";
            this.lblFileExcel.Size = new System.Drawing.Size(0, 13);
            this.lblFileExcel.TabIndex = 1;
            // 
            // opdFiles
            // 
            this.opdFiles.FileName = "openFileDialog1";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fichierToolStripMenuItem,
            this.aideToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(879, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fichierToolStripMenuItem
            // 
            this.fichierToolStripMenuItem.Name = "fichierToolStripMenuItem";
            this.fichierToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.fichierToolStripMenuItem.Text = "Fichier";
            // 
            // aideToolStripMenuItem
            // 
            this.aideToolStripMenuItem.Name = "aideToolStripMenuItem";
            this.aideToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.aideToolStripMenuItem.Text = "Aide";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(879, 313);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Replication Excel";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lsbNames;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lblFileExcel;
        private System.Windows.Forms.OpenFileDialog opdFiles;
        private System.Windows.Forms.TextBox tbxAddName;
        private System.Windows.Forms.Button btnAddName;
        private System.Windows.Forms.Button btnSaveName;
        private System.Windows.Forms.Button btnCopyExcel;
        private System.Windows.Forms.Button btnDeleteName;
        private System.Windows.Forms.ListBox lsbSheets;
        private System.Windows.Forms.SaveFileDialog sfdFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxEncoding;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBrowseNomsEleves;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbxExcelFile;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fichierToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aideToolStripMenuItem;
        private System.Windows.Forms.Label label9;
    }
}

