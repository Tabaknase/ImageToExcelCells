namespace PictureToExcelCell
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
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.btnChoosePicture = new System.Windows.Forms.Button();
            this.chkWithMakro = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(12, 12);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(120, 20);
            this.numericUpDown1.TabIndex = 0;
            this.numericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numericUpDown1.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            // 
            // btnChoosePicture
            // 
            this.btnChoosePicture.Location = new System.Drawing.Point(12, 38);
            this.btnChoosePicture.Name = "btnChoosePicture";
            this.btnChoosePicture.Size = new System.Drawing.Size(120, 23);
            this.btnChoosePicture.TabIndex = 1;
            this.btnChoosePicture.Text = "Choose picture";
            this.btnChoosePicture.UseVisualStyleBackColor = true;
            this.btnChoosePicture.Click += new System.EventHandler(this.btnChoosePicture_Click);
            // 
            // chkWithMakro
            // 
            this.chkWithMakro.AutoSize = true;
            this.chkWithMakro.Location = new System.Drawing.Point(12, 67);
            this.chkWithMakro.Name = "chkWithMakro";
            this.chkWithMakro.Size = new System.Drawing.Size(87, 17);
            this.chkWithMakro.TabIndex = 2;
            this.chkWithMakro.Text = "With Makkro";
            this.chkWithMakro.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(147, 103);
            this.Controls.Add(this.chkWithMakro);
            this.Controls.Add(this.btnChoosePicture);
            this.Controls.Add(this.numericUpDown1);
            this.Name = "Form1";
            this.Text = "Form";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Button btnChoosePicture;
        private System.Windows.Forms.CheckBox chkWithMakro;
    }
}

