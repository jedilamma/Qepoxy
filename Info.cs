using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Qepoxy
{
    public partial class Info : Form
    {
        public Info()
        {
            InitializeComponent();
            Copywritelbl.Text = "©2022 Copyright Jeffrey Whitford\n" +
                "Email: Jeff.Whitford@outlook.com.au\n" +
                "Ph: 0415 591 321\n\n" +
                "Licensed to: Jacob Young - Qepoxy\n" +
                "1/25 Depot st Banyo QLD 4014\n" +
                "Email: Jacob@Qepoxy.com.au\n" +
                "Ph: 0499 554 505";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
