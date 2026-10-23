
namespace Qepoxy
{
    partial class StartForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StartForm));
            this.Loginbtn = new System.Windows.Forms.Button();
            this.Passwordtb = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.ConnectBtn = new System.Windows.Forms.Button();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.ServerConnectCb = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.DcBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.Statuslbl = new System.Windows.Forms.Label();
            this.Usernametb = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // Loginbtn
            // 
            this.Loginbtn.Enabled = false;
            this.Loginbtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.Loginbtn.Location = new System.Drawing.Point(7, 14);
            this.Loginbtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Loginbtn.Name = "Loginbtn";
            this.Loginbtn.Size = new System.Drawing.Size(110, 53);
            this.Loginbtn.TabIndex = 4;
            this.Loginbtn.Text = "Login";
            this.Loginbtn.UseVisualStyleBackColor = true;
            this.Loginbtn.Click += new System.EventHandler(this.Loginbtn_Click);
            // 
            // Passwordtb
            // 
            this.Passwordtb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Passwordtb.Location = new System.Drawing.Point(139, 71);
            this.Passwordtb.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Passwordtb.Name = "Passwordtb";
            this.Passwordtb.Size = new System.Drawing.Size(194, 23);
            this.Passwordtb.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(139, 6);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 15);
            this.label1.TabIndex = 106;
            this.label1.Text = "Username";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(139, 53);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 15);
            this.label2.TabIndex = 107;
            this.label2.Text = "Password";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(392, 151);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 15);
            this.label3.TabIndex = 108;
            this.label3.Text = "Server Status:";
            // 
            // ConnectBtn
            // 
            this.ConnectBtn.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ConnectBtn.Location = new System.Drawing.Point(521, 145);
            this.ConnectBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.ConnectBtn.Name = "ConnectBtn";
            this.ConnectBtn.Size = new System.Drawing.Size(85, 27);
            this.ConnectBtn.TabIndex = 5;
            this.ConnectBtn.Text = "Connect";
            this.ConnectBtn.UseVisualStyleBackColor = true;
            this.ConnectBtn.Click += new System.EventHandler(this.Connect1_Click);
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::Qepoxy.Properties.Resources.red1;
            this.pictureBox2.Location = new System.Drawing.Point(477, 145);
            this.pictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(36, 35);
            this.pictureBox2.TabIndex = 110;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.ErrorImage = global::Qepoxy.Properties.Resources.Qepoxy_sml;
            this.pictureBox1.Image = global::Qepoxy.Properties.Resources.Qepoxy_sml;
            this.pictureBox1.InitialImage = global::Qepoxy.Properties.Resources.Qepoxy_sml;
            this.pictureBox1.Location = new System.Drawing.Point(376, -13);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(321, 133);
            this.pictureBox1.TabIndex = 14;
            this.pictureBox1.TabStop = false;
            // 
            // ServerConnectCb
            // 
            this.ServerConnectCb.FormattingEnabled = true;
            this.ServerConnectCb.Location = new System.Drawing.Point(139, 145);
            this.ServerConnectCb.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.ServerConnectCb.Name = "ServerConnectCb";
            this.ServerConnectCb.Size = new System.Drawing.Size(229, 23);
            this.ServerConnectCb.TabIndex = 3;
            this.ServerConnectCb.SelectedIndexChanged += new System.EventHandler(this.ServerConnectCb_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(92, 148);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(39, 15);
            this.label4.TabIndex = 107;
            this.label4.Text = "Server";
            // 
            // DcBtn
            // 
            this.DcBtn.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.DcBtn.Location = new System.Drawing.Point(609, 145);
            this.DcBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.DcBtn.Name = "DcBtn";
            this.DcBtn.Size = new System.Drawing.Size(85, 27);
            this.DcBtn.TabIndex = 6;
            this.DcBtn.Text = "Disconnect";
            this.DcBtn.UseVisualStyleBackColor = true;
            this.DcBtn.Click += new System.EventHandler(this.DcBtn_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button1.Location = new System.Drawing.Point(338, 69);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(27, 29);
            this.button1.TabIndex = 112;
            this.button1.Text = "👁";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button1_MouseDown);
            this.button1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.button1_MouseUp);
            // 
            // Statuslbl
            // 
            this.Statuslbl.AutoSize = true;
            this.Statuslbl.Location = new System.Drawing.Point(139, 101);
            this.Statuslbl.Name = "Statuslbl";
            this.Statuslbl.Size = new System.Drawing.Size(0, 15);
            this.Statuslbl.TabIndex = 113;
            // 
            // Usernametb
            // 
            this.Usernametb.Location = new System.Drawing.Point(139, 25);
            this.Usernametb.Name = "Usernametb";
            this.Usernametb.Size = new System.Drawing.Size(194, 23);
            this.Usernametb.TabIndex = 1;
            // 
            // StartForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(707, 192);
            this.Controls.Add(this.Usernametb);
            this.Controls.Add(this.Statuslbl);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.ServerConnectCb);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.DcBtn);
            this.Controls.Add(this.ConnectBtn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Passwordtb);
            this.Controls.Add(this.Loginbtn);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "StartForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Qepoxy";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button Loginbtn;
        private System.Windows.Forms.TextBox Passwordtb;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button ConnectBtn;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.ComboBox ServerConnectCb;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button DcBtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label Statuslbl;
        private System.Windows.Forms.TextBox Usernametb;
    }
}