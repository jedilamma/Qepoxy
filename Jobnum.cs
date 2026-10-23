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
    public partial class Jobnum : Form
    {
        public Jobnum()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            string day = DateTime.Now.ToString("dd");
            string month = DateTime.Now.ToString("MM");
            string year = DateTime.Now.ToString("yy");
            string hour = DateTime.Now.ToString("HH");
            string minute = DateTime.Now.ToString("mm");
            string second = DateTime.Now.ToString("ss");

            string out1 = genCode(day);
            string out2 = genCode(month);
            string out3 = genCode(year);
            string out4 = genCode(hour);
            string out5 = genCode(minute);
            string out6 = genCode(second);

            //string final = month + day + year;// +"-"+ out1 + out2 + out3 + out4 + out5+ out6;
            string final = month + day + out6 + out5 + year;// +"-"+ out1 + out2 + out3 + out4 + out5+ out6;
            //string final = out1 + out2 + out3 + out4 + out5 + out6;
            GenDispLabel.Text = "QE"+final;
            //byte[] ba = Encoding.Default.GetBytes(year+hour+minute);
            //var hexString = BitConverter.ToString(ba);
            //hexString = hexString.Replace("-", "");

            //GenDispLabel
        }

        private string genCode(string input)
        {
            string endval = "";
            int val = int.Parse(input);
            switch (val)
            {
                case 1:
                    endval = "1";
                    break;
                case 2:
                    endval = "2";
                    break;
                case 3:
                    endval = "3";
                    break;
                case 4:
                    endval = "4";
                    break;
                case 5:
                    endval = "5";
                    break;
                case 6:
                    endval = "6";
                    break;
                case 7:
                    endval = "7";
                    break;
                case 8:
                    endval = "8";
                    break;
                case 9:
                    endval = "9";
                    break;
                case 10:
                    endval = "a";
                    break;
                case 11:
                    endval = "b";
                    break;
                case 12:
                    endval = "c";
                    break;
                case 13:
                    endval = "d";
                    break;
                case 14:
                    endval = "e";
                    break;
                case 15:
                    endval = "f";
                    break;
                case 16:
                    endval = "g";
                    break;
                case 17:
                    endval = "h";
                    break;
                case 18:
                    endval = "i";
                    break;
                case 19:
                    endval = "j";
                    break;
                case 20:
                    endval = "k";
                    break;
                case 21:
                    endval = "l";
                    break;
                case 22:
                    endval = "m";
                    break;
                case 23:
                    endval = "n";
                    break;
                case 24:
                    endval = "o";
                    break;
                case 25:
                    endval = "p";
                    break;
                case 26:
                    endval = "q";
                    break;
                case 27:
                    endval = "r";
                    break;
                case 28:
                    endval = "s";
                    break;
                case 29:
                    endval = "t";
                    break;
                case 30:
                    endval = "u";
                    break;
                case 31:
                    endval = "v";
                    break;
                case 32:
                    endval = "w";
                    break;
                case 33:
                    endval = "x";
                    break;
                case 34:
                    endval = "y";
                    break;
                case 35:
                    endval = "z";
                    break;
                case 36:
                    endval = "A";
                    break;
                case 37:
                    endval = "B";
                    break;
                case 38:
                    endval = "C";
                    break;
                case 39:
                    endval = "D";
                    break;
                case 40:
                    endval = "E";
                    break;
                case 41:
                    endval = "F";
                    break;
                case 42:
                    endval = "G";
                    break;
                case 43:
                    endval = "H";
                    break;
                case 44:
                    endval = "I";
                    break;
                case 45:
                    endval = "J";
                    break;
                case 46:
                    endval = "K";
                    break;
                case 47:
                    endval = "L";
                    break;
                case 48:
                    endval = "M";
                    break;
                case 49:
                    endval = "N";
                    break;
                case 50:
                    endval = "O";
                    break;
                case 51:
                    endval = "P";
                    break;
                case 52:
                    endval = "Q";
                    break;
                case 53:
                    endval = "R";
                    break;
                case 54:
                    endval = "S";
                    break;
                case 55:
                    endval = "1";
                    break;
                case 56:
                    endval = "T";
                    break;
                case 57:
                    endval = "U";
                    break;
                case 58:
                    endval = "V";
                    break;
                case 59:
                    endval = "W";
                    break;
                case 60:
                    endval = "X";
                    break;
                case 0:
                    endval = "Y";
                    break;
            }
            return endval;
        }
    }
}
