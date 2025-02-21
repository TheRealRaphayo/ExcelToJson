using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ExcelToJsonApp
{
    public partial class Form1 : Form
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
            this.BtnSelectFile = new System.Windows.Forms.Button();
            this.TextFilePath = new System.Windows.Forms.TextBox();
            this.BtnExportJson = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BtnSelectFile
            // 
            this.BtnSelectFile.Location = new System.Drawing.Point(12, 12);
            this.BtnSelectFile.Name = "BtnSelectFile";
            this.BtnSelectFile.Size = new System.Drawing.Size(566, 50);
            this.BtnSelectFile.TabIndex = 0;
            this.BtnSelectFile.Text = "Select Excel File";
            this.BtnSelectFile.UseVisualStyleBackColor = true;
            this.BtnSelectFile.Click += new System.EventHandler(this.BtnSelectFile_Click);
            // 
            // TextFilePath
            // 
            this.TextFilePath.Location = new System.Drawing.Point(12, 68);
            this.TextFilePath.Name = "TextFilePath";
            this.TextFilePath.ReadOnly = true;
            this.TextFilePath.Size = new System.Drawing.Size(566, 20);
            this.TextFilePath.TabIndex = 1;
            this.TextFilePath.TextChanged += new System.EventHandler(this.TextFilePath_TextChanged);
            // 
            // BtnExportJson
            // 
            this.BtnExportJson.Location = new System.Drawing.Point(12, 103);
            this.BtnExportJson.Name = "BtnExportJson";
            this.BtnExportJson.Size = new System.Drawing.Size(565, 62);
            this.BtnExportJson.TabIndex = 2;
            this.BtnExportJson.Text = "Export to JSON";
            this.BtnExportJson.UseVisualStyleBackColor = true;
            this.BtnExportJson.Click += new System.EventHandler(this.BtnExportJson_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 181);
            this.Controls.Add(this.BtnExportJson);
            this.Controls.Add(this.TextFilePath);
            this.Controls.Add(this.BtnSelectFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }



        #endregion

        private Button BtnSelectFile;
        private TextBox TextFilePath;
        private Button BtnExportJson;
    }
}

