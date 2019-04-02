namespace WindowsFormsApp1
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
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(201, 66);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(159, 76);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(201, 164);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(159, 50);
            this.button2.TabIndex = 1;
            this.button2.Text = "Excel to CSV ";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(201, 248);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(159, 53);
            this.button3.TabIndex = 2;
            this.button3.Text = "CSV To Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(462, 255);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(159, 46);
            this.button4.TabIndex = 3;
            this.button4.Text = "Excel_Delete_BlankColumns";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // btnGetHeaderClmn
            // 
            this.btnGetHeaderClmn.Location = new System.Drawing.Point(456, 169);
            this.btnGetHeaderClmn.Name = "btnGetHeaderClmn";
            this.btnGetHeaderClmn.Size = new System.Drawing.Size(194, 57);
            this.btnGetHeaderClmn.TabIndex = 4;
            this.btnGetHeaderClmn.Text = "Excel Get Header Column Number";
            this.btnGetHeaderClmn.UseVisualStyleBackColor = true;
            this.btnGetHeaderClmn.Click += new System.EventHandler(this.btnGetHeaderClmn_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(462, 76);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(188, 66);
            this.button5.TabIndex = 5;
            this.button5.Text = "Test Run Macro";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCreateNewExcel
            // 
            this.btnCreateNewExcel.Location = new System.Drawing.Point(627, 251);
            this.btnCreateNewExcel.Name = "btnCreateNewExcel";
            this.btnCreateNewExcel.Size = new System.Drawing.Size(144, 54);
            this.btnCreateNewExcel.TabIndex = 6;
            this.btnCreateNewExcel.Text = "Excel create New Excel";
            this.btnCreateNewExcel.UseVisualStyleBackColor = true;
            this.btnCreateNewExcel.Click += new System.EventHandler(this.btnCreateNewExcel_Click);
            // 
            // btnDeleteRow
            // 
            this.btnDeleteRow.Location = new System.Drawing.Point(21, 306);
            this.btnDeleteRow.Name = "btnDeleteRow";
            this.btnDeleteRow.Size = new System.Drawing.Size(136, 85);
            this.btnDeleteRow.TabIndex = 7;
            this.btnDeleteRow.Text = "Excel Delete Row";
            this.btnDeleteRow.UseVisualStyleBackColor = true;
            this.btnDeleteRow.Click += new System.EventHandler(this.btnDeleteRow_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(21, 98);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(136, 58);
            this.button6.TabIndex = 8;
            this.button6.Text = "Excel Copy Data";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(201, 322);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(159, 53);
            this.button7.TabIndex = 9;
            this.button7.Text = "Htm To Excel";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(462, 329);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(159, 46);
            this.button8.TabIndex = 10;
            this.button8.Text = "Excel_Delete_Duplicate Headings";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
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
    }
}

