
namespace Qepoxy
{
    partial class DataControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DataControl));
            this.button3 = new System.Windows.Forms.Button();
            this.SQLerrorlv = new System.Windows.Forms.ListView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Legacycb = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.DataLoadtb = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Clearbtn = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.Datasavelocationtb = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.DataPG = new System.Windows.Forms.ProgressBar();
            this.Displaylbl = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // button3
            // 
            resources.ApplyResources(this.button3, "button3");
            this.button3.Name = "button3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // SQLerrorlv
            // 
            this.SQLerrorlv.HideSelection = false;
            resources.ApplyResources(this.SQLerrorlv, "SQLerrorlv");
            this.SQLerrorlv.Name = "SQLerrorlv";
            this.SQLerrorlv.UseCompatibleStateImageBehavior = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Legacycb);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.DataLoadtb);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.button3);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // Legacycb
            // 
            resources.ApplyResources(this.Legacycb, "Legacycb");
            this.Legacycb.Name = "Legacycb";
            this.Legacycb.UseVisualStyleBackColor = true;
            this.Legacycb.CheckedChanged += new System.EventHandler(this.Legacycb_CheckedChanged);
            // 
            // button1
            // 
            this.button1.BackgroundImage = global::Qepoxy.Properties.Resources.open_file_icon_sml;
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // DataLoadtb
            // 
            resources.ApplyResources(this.DataLoadtb, "DataLoadtb");
            this.DataLoadtb.Name = "DataLoadtb";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // Clearbtn
            // 
            resources.ApplyResources(this.Clearbtn, "Clearbtn");
            this.Clearbtn.Name = "Clearbtn";
            this.Clearbtn.UseVisualStyleBackColor = true;
            this.Clearbtn.Click += new System.EventHandler(this.Clearbtn_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button6);
            this.groupBox2.Controls.Add(this.Datasavelocationtb);
            this.groupBox2.Controls.Add(this.label1);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // button2
            // 
            this.button2.BackgroundImage = global::Qepoxy.Properties.Resources.open_file_icon_sml;
            resources.ApplyResources(this.button2, "button2");
            this.button2.Name = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button6
            // 
            resources.ApplyResources(this.button6, "button6");
            this.button6.Name = "button6";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // Datasavelocationtb
            // 
            resources.ApplyResources(this.Datasavelocationtb, "Datasavelocationtb");
            this.Datasavelocationtb.Name = "Datasavelocationtb";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // DataPG
            // 
            resources.ApplyResources(this.DataPG, "DataPG");
            this.DataPG.Name = "DataPG";
            // 
            // Displaylbl
            // 
            resources.ApplyResources(this.Displaylbl, "Displaylbl");
            this.Displaylbl.Name = "Displaylbl";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Qepoxy.Properties.Resources.Qepoxy_sml2;
            resources.ApplyResources(this.pictureBox1, "pictureBox1");
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.TabStop = false;
            // 
            // DataControl
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.Displaylbl);
            this.Controls.Add(this.DataPG);
            this.Controls.Add(this.Clearbtn);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.SQLerrorlv);
            this.Controls.Add(this.pictureBox1);
            this.MaximizeBox = false;
            this.Name = "DataControl";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ListView SQLerrorlv;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TextBox Datasavelocationtb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox DataLoadtb;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button Clearbtn;
        private System.Windows.Forms.CheckBox Legacycb;
        private System.Windows.Forms.ProgressBar DataPG;
        private System.Windows.Forms.Label Displaylbl;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}