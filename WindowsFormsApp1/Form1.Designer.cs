﻿namespace WindowsFormsApp1
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.btnGetHeaderClmn = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.btnCreateNewExcel = new System.Windows.Forms.Button();
            this.btnDeleteRow = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(268, 76);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(212, 88);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(268, 189);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(212, 58);
            this.button2.TabIndex = 1;
            this.button2.Text = "Excel to CSV ";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(268, 286);
            this.button3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(212, 61);
            this.button3.TabIndex = 2;
            this.button3.Text = "CSV To Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(616, 294);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(212, 53);
            this.button4.TabIndex = 3;
            this.button4.Text = "Excel_Delete_BlankColumns";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // btnGetHeaderClmn
            // 
            this.btnGetHeaderClmn.Location = new System.Drawing.Point(608, 195);
            this.btnGetHeaderClmn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnGetHeaderClmn.Name = "btnGetHeaderClmn";
            this.btnGetHeaderClmn.Size = new System.Drawing.Size(259, 66);
            this.btnGetHeaderClmn.TabIndex = 4;
            this.btnGetHeaderClmn.Text = "Excel Get Header Column Number";
            this.btnGetHeaderClmn.UseVisualStyleBackColor = true;
            this.btnGetHeaderClmn.Click += new System.EventHandler(this.btnGetHeaderClmn_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(616, 88);
            this.button5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(251, 76);
            this.button5.TabIndex = 5;
            this.button5.Text = "Test Run Macro";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCreateNewExcel
            // 
            this.btnCreateNewExcel.Location = new System.Drawing.Point(836, 290);
            this.btnCreateNewExcel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnCreateNewExcel.Name = "btnCreateNewExcel";
            this.btnCreateNewExcel.Size = new System.Drawing.Size(192, 62);
            this.btnCreateNewExcel.TabIndex = 6;
            this.btnCreateNewExcel.Text = "Excel create New Excel";
            this.btnCreateNewExcel.UseVisualStyleBackColor = true;
            this.btnCreateNewExcel.Click += new System.EventHandler(this.btnCreateNewExcel_Click);
            // 
            // btnDeleteRow
            // 
            this.btnDeleteRow.Location = new System.Drawing.Point(28, 353);
            this.btnDeleteRow.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnDeleteRow.Name = "btnDeleteRow";
            this.btnDeleteRow.Size = new System.Drawing.Size(181, 98);
            this.btnDeleteRow.TabIndex = 7;
            this.btnDeleteRow.Text = "Excel Delete Row";
            this.btnDeleteRow.UseVisualStyleBackColor = true;
            this.btnDeleteRow.Click += new System.EventHandler(this.btnDeleteRow_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(28, 113);
            this.button6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(181, 67);
            this.button6.TabIndex = 8;
            this.button6.Text = "Excel Copy Data";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(268, 372);
            this.button7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(212, 61);
            this.button7.TabIndex = 9;
            this.button7.Text = "Htm To Excel";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(616, 380);
            this.button8.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(212, 53);
            this.button8.TabIndex = 10;
            this.button8.Text = "Excel_Delete_Duplicate Headings";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(28, 237);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(181, 76);
            this.button9.TabIndex = 11;
            this.button9.Text = "Excel_Walmart_Reconciliation";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(616, 456);
            this.button10.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(212, 50);
            this.button10.TabIndex = 12;
            this.button10.Text = "Remove Duplicates";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(842, 380);
            this.button11.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(186, 53);
            this.button11.TabIndex = 13;
            this.button11.Text = "Excel_Filter_Delete_Row";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 519);
            this.Controls.Add(this.button11);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.btnDeleteRow);
            this.Controls.Add(this.btnCreateNewExcel);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.btnGetHeaderClmn);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button btnGetHeaderClmn;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button btnCreateNewExcel;
        private System.Windows.Forms.Button btnDeleteRow;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
    }
}

