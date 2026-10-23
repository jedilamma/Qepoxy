namespace Qepoxy
{
    partial class POItemEdit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(POItemEdit));
            this.Savebtn = new System.Windows.Forms.Button();
            this.Cancelbtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.CostPtb = new System.Windows.Forms.TextBox();
            this.Qtytb = new System.Windows.Forms.TextBox();
            this.Desctb = new System.Windows.Forms.TextBox();
            this.Pricetb = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Calc = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.Packamnttb = new System.Windows.Forms.TextBox();
            this.Unitcb = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // Savebtn
            // 
            this.Savebtn.Location = new System.Drawing.Point(13, 13);
            this.Savebtn.Name = "Savebtn";
            this.Savebtn.Size = new System.Drawing.Size(75, 23);
            this.Savebtn.TabIndex = 0;
            this.Savebtn.Text = "Save";
            this.Savebtn.UseVisualStyleBackColor = true;
            this.Savebtn.Click += new System.EventHandler(this.Savebtn_Click);
            // 
            // Cancelbtn
            // 
            this.Cancelbtn.Location = new System.Drawing.Point(95, 13);
            this.Cancelbtn.Name = "Cancelbtn";
            this.Cancelbtn.Size = new System.Drawing.Size(75, 23);
            this.Cancelbtn.TabIndex = 1;
            this.Cancelbtn.Text = "Cancel";
            this.Cancelbtn.UseVisualStyleBackColor = true;
            this.Cancelbtn.Click += new System.EventHandler(this.Cancelbtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Description";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Quantity";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Cost Per";
            // 
            // CostPtb
            // 
            this.CostPtb.Location = new System.Drawing.Point(70, 127);
            this.CostPtb.Name = "CostPtb";
            this.CostPtb.Size = new System.Drawing.Size(135, 20);
            this.CostPtb.TabIndex = 5;
            // 
            // Qtytb
            // 
            this.Qtytb.Location = new System.Drawing.Point(70, 74);
            this.Qtytb.Name = "Qtytb";
            this.Qtytb.Size = new System.Drawing.Size(135, 20);
            this.Qtytb.TabIndex = 6;
            // 
            // Desctb
            // 
            this.Desctb.Location = new System.Drawing.Point(70, 48);
            this.Desctb.Name = "Desctb";
            this.Desctb.Size = new System.Drawing.Size(135, 20);
            this.Desctb.TabIndex = 7;
            // 
            // Pricetb
            // 
            this.Pricetb.Location = new System.Drawing.Point(70, 153);
            this.Pricetb.Name = "Pricetb";
            this.Pricetb.Size = new System.Drawing.Size(135, 20);
            this.Pricetb.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 156);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Price";
            // 
            // Calc
            // 
            this.Calc.Location = new System.Drawing.Point(130, 179);
            this.Calc.Name = "Calc";
            this.Calc.Size = new System.Drawing.Size(75, 23);
            this.Calc.TabIndex = 10;
            this.Calc.Text = "Calculate";
            this.Calc.UseVisualStyleBackColor = true;
            this.Calc.Click += new System.EventHandler(this.Calc_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 104);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Pack Amnt";
            // 
            // Packamnttb
            // 
            this.Packamnttb.Location = new System.Drawing.Point(70, 101);
            this.Packamnttb.Name = "Packamnttb";
            this.Packamnttb.Size = new System.Drawing.Size(79, 20);
            this.Packamnttb.TabIndex = 12;
            // 
            // Unitcb
            // 
            this.Unitcb.FormattingEnabled = true;
            this.Unitcb.Location = new System.Drawing.Point(156, 101);
            this.Unitcb.Name = "Unitcb";
            this.Unitcb.Size = new System.Drawing.Size(49, 21);
            this.Unitcb.TabIndex = 13;
            // 
            // POItemEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(217, 216);
            this.Controls.Add(this.Unitcb);
            this.Controls.Add(this.Packamnttb);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Calc);
            this.Controls.Add(this.Pricetb);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.Desctb);
            this.Controls.Add(this.Qtytb);
            this.Controls.Add(this.CostPtb);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Cancelbtn);
            this.Controls.Add(this.Savebtn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "POItemEdit";
            this.Text = "Purchase Order Item Editor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Savebtn;
        private System.Windows.Forms.Button Cancelbtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox CostPtb;
        private System.Windows.Forms.TextBox Qtytb;
        private System.Windows.Forms.TextBox Desctb;
        private System.Windows.Forms.TextBox Pricetb;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button Calc;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Packamnttb;
        private System.Windows.Forms.ComboBox Unitcb;
    }
}