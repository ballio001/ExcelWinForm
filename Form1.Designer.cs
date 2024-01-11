
namespace ExcelWinForm
{
    partial class Form1
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
            this.CmdRead = new System.Windows.Forms.Button();
            this.CmdWrite = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CmdRead
            // 
            this.CmdRead.Location = new System.Drawing.Point(12, 34);
            this.CmdRead.Name = "CmdRead";
            this.CmdRead.Size = new System.Drawing.Size(75, 23);
            this.CmdRead.TabIndex = 0;
            this.CmdRead.Text = "Read";
            this.CmdRead.UseVisualStyleBackColor = true;
            this.CmdRead.Click += new System.EventHandler(this.CmdRead_Click);
            // 
            // CmdWrite
            // 
            //this.CmdWrite.Location = new System.Drawing.Point(102, 34);
            //this.CmdWrite.Name = "CmdWrite";
            //this.CmdWrite.Size = new System.Drawing.Size(75, 23);
            //this.CmdWrite.TabIndex = 1;
            //this.CmdWrite.Text = "Write";
            //this.CmdWrite.UseVisualStyleBackColor = true;
            //this.CmdWrite.Click += new System.EventHandler(this.CmdWrite_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 431);
            this.Controls.Add(this.CmdWrite);
            this.Controls.Add(this.CmdRead);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button CmdRead;
        private System.Windows.Forms.Button CmdWrite;
    }
}