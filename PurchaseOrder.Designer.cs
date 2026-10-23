namespace Qepoxy
{
    partial class PurchaseOrder
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
            this.POgenbtn = new System.Windows.Forms.Button();
            this.PoCostingslv = new System.Windows.Forms.ListView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.Attentb = new System.Windows.Forms.TextBox();
            this.Companytb = new System.Windows.Forms.TextBox();
            this.Emailtb = new System.Windows.Forms.TextBox();
            this.Deliverytb = new System.Windows.Forms.TextBox();
            this.PoNumtb = new System.Windows.Forms.TextBox();
            this.Datetb = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.Addbtn = new System.Windows.Forms.Button();
            this.Editbtn = new System.Windows.Forms.Button();
            this.Rembtn = new System.Windows.Forms.Button();
            this.Productlv = new System.Windows.Forms.ListView();
            this.label8 = new System.Windows.Forms.Label();
            this.Arealb = new System.Windows.Forms.Label();
            this.Quantitytb = new System.Windows.Forms.TextBox();
            this.Unitcb = new System.Windows.Forms.ComboBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // POgenbtn
            // 
            this.POgenbtn.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.POgenbtn.Location = new System.Drawing.Point(12, 12);
            this.POgenbtn.Name = "POgenbtn";
            this.POgenbtn.Size = new System.Drawing.Size(173, 68);
            this.POgenbtn.TabIndex = 0;
            this.POgenbtn.Text = "Generate Purchase Order";
            this.POgenbtn.UseVisualStyleBackColor = true;
            this.POgenbtn.Click += new System.EventHandler(this.POgenbtn_Click);
            // 
            // PoCostingslv
            // 
            this.PoCostingslv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.PoCostingslv.HideSelection = false;
            this.PoCostingslv.Location = new System.Drawing.Point(12, 119);
            this.PoCostingslv.Name = "PoCostingslv";
            this.PoCostingslv.Size = new System.Drawing.Size(428, 578);
            this.PoCostingslv.TabIndex = 1;
            this.PoCostingslv.UseCompatibleStateImageBehavior = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 103);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Order Item List";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(204, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Attention:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(202, 52);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Company:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(221, 78);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Email:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(456, 24);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(33, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Date:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(424, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "PO Number:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(444, 78);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 8;
            this.label7.Text = "Delivery";
            // 
            // Attentb
            // 
            this.Attentb.Location = new System.Drawing.Point(262, 21);
            this.Attentb.Name = "Attentb";
            this.Attentb.Size = new System.Drawing.Size(133, 20);
            this.Attentb.TabIndex = 9;
            // 
            // Companytb
            // 
            this.Companytb.Location = new System.Drawing.Point(262, 49);
            this.Companytb.Name = "Companytb";
            this.Companytb.Size = new System.Drawing.Size(133, 20);
            this.Companytb.TabIndex = 10;
            // 
            // Emailtb
            // 
            this.Emailtb.Location = new System.Drawing.Point(262, 75);
            this.Emailtb.Name = "Emailtb";
            this.Emailtb.Size = new System.Drawing.Size(133, 20);
            this.Emailtb.TabIndex = 11;
            // 
            // Deliverytb
            // 
            this.Deliverytb.Location = new System.Drawing.Point(495, 75);
            this.Deliverytb.Name = "Deliverytb";
            this.Deliverytb.Size = new System.Drawing.Size(133, 20);
            this.Deliverytb.TabIndex = 14;
            // 
            // PoNumtb
            // 
            this.PoNumtb.Location = new System.Drawing.Point(495, 49);
            this.PoNumtb.Name = "PoNumtb";
            this.PoNumtb.Size = new System.Drawing.Size(133, 20);
            this.PoNumtb.TabIndex = 13;
            // 
            // Datetb
            // 
            this.Datetb.Location = new System.Drawing.Point(495, 21);
            this.Datetb.Name = "Datetb";
            this.Datetb.Size = new System.Drawing.Size(133, 20);
            this.Datetb.TabIndex = 12;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(634, 19);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(55, 22);
            this.button2.TabIndex = 40;
            this.button2.Text = "Date";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Addbtn
            // 
            this.Addbtn.Location = new System.Drawing.Point(446, 188);
            this.Addbtn.Name = "Addbtn";
            this.Addbtn.Size = new System.Drawing.Size(75, 23);
            this.Addbtn.TabIndex = 41;
            this.Addbtn.Text = "Add";
            this.Addbtn.UseVisualStyleBackColor = true;
            this.Addbtn.Click += new System.EventHandler(this.Addbtn_Click);
            // 
            // Editbtn
            // 
            this.Editbtn.Location = new System.Drawing.Point(446, 218);
            this.Editbtn.Name = "Editbtn";
            this.Editbtn.Size = new System.Drawing.Size(75, 23);
            this.Editbtn.TabIndex = 42;
            this.Editbtn.Text = "Edit";
            this.Editbtn.UseVisualStyleBackColor = true;
            this.Editbtn.Click += new System.EventHandler(this.Editbtn_Click);
            // 
            // Rembtn
            // 
            this.Rembtn.Location = new System.Drawing.Point(446, 248);
            this.Rembtn.Name = "Rembtn";
            this.Rembtn.Size = new System.Drawing.Size(75, 23);
            this.Rembtn.TabIndex = 43;
            this.Rembtn.Text = "Remove";
            this.Rembtn.UseVisualStyleBackColor = true;
            this.Rembtn.Click += new System.EventHandler(this.Rembtn_Click);
            // 
            // Productlv
            // 
            this.Productlv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.Productlv.HideSelection = false;
            this.Productlv.Location = new System.Drawing.Point(527, 169);
            this.Productlv.Name = "Productlv";
            this.Productlv.Size = new System.Drawing.Size(452, 528);
            this.Productlv.TabIndex = 44;
            this.Productlv.UseCompatibleStateImageBehavior = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(527, 153);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 13);
            this.label8.TabIndex = 45;
            this.label8.Text = "Products";
            // 
            // Arealb
            // 
            this.Arealb.AutoSize = true;
            this.Arealb.Location = new System.Drawing.Point(784, 146);
            this.Arealb.Name = "Arealb";
            this.Arealb.Size = new System.Drawing.Size(23, 13);
            this.Arealb.TabIndex = 46;
            this.Arealb.Text = "Qty";
            // 
            // Quantitytb
            // 
            this.Quantitytb.Location = new System.Drawing.Point(813, 143);
            this.Quantitytb.Name = "Quantitytb";
            this.Quantitytb.Size = new System.Drawing.Size(105, 20);
            this.Quantitytb.TabIndex = 47;
            // 
            // Unitcb
            // 
            this.Unitcb.FormattingEnabled = true;
            this.Unitcb.Location = new System.Drawing.Point(924, 142);
            this.Unitcb.Name = "Unitcb";
            this.Unitcb.Size = new System.Drawing.Size(55, 21);
            this.Unitcb.TabIndex = 48;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Qepoxy.Properties.Resources.Qepoxy_sml2;
            this.pictureBox1.Location = new System.Drawing.Point(740, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(239, 104);
            this.pictureBox1.TabIndex = 49;
            this.pictureBox1.TabStop = false;
            // 
            // PurchaseOrder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(991, 709);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Unitcb);
            this.Controls.Add(this.Quantitytb);
            this.Controls.Add(this.Arealb);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.Productlv);
            this.Controls.Add(this.Rembtn);
            this.Controls.Add(this.Editbtn);
            this.Controls.Add(this.Addbtn);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.Deliverytb);
            this.Controls.Add(this.PoNumtb);
            this.Controls.Add(this.Datetb);
            this.Controls.Add(this.Emailtb);
            this.Controls.Add(this.Companytb);
            this.Controls.Add(this.Attentb);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PoCostingslv);
            this.Controls.Add(this.POgenbtn);
            this.Name = "PurchaseOrder";
            this.Text = "Purchase Order";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button POgenbtn;
        private System.Windows.Forms.ListView PoCostingslv;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox Attentb;
        private System.Windows.Forms.TextBox Companytb;
        private System.Windows.Forms.TextBox Emailtb;
        private System.Windows.Forms.TextBox Deliverytb;
        private System.Windows.Forms.TextBox PoNumtb;
        private System.Windows.Forms.TextBox Datetb;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button Addbtn;
        private System.Windows.Forms.Button Editbtn;
        private System.Windows.Forms.Button Rembtn;
        private System.Windows.Forms.ListView Productlv;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label Arealb;
        private System.Windows.Forms.TextBox Quantitytb;
        private System.Windows.Forms.ComboBox Unitcb;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}