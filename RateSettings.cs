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
    public partial class hrAddbtn : Form
    {
        Backend newBe;
        public hrAddbtn(Backend Backend)
        {
            InitializeComponent();
            newBe = Backend;

            this.Ratelistlv.Scrollable = true;
            this.Ratelistlv.View = View.Details;
            this.Ratelistlv.GridLines = true;
            this.Ratelistlv.FullRowSelect = true;
            this.Ratelistlv.Columns.Add("Method", 220);
            this.Ratelistlv.Columns.Add("Rate", 55);

            this.Addlv.Scrollable = true;
            this.Addlv.View = View.Details;
            this.Addlv.GridLines = true;
            this.Addlv.FullRowSelect = true;
            this.Addlv.Columns.Add("ItemName", 320);

            populatAdditions();
            populateHrrates();

            this.FormClosing += Form_FormClosing;
        }
        private void Form_FormClosing(Object sender, FormClosingEventArgs e)
        {
            //foreach (Hrrates hritem in newBe.lists.HrrateList)
            //{
            //    newBe.lists.UpdateHrrates(hritem);
            //}

            Main form1 = (Main)this.Owner;
            form1.Populate();
            form1.populatAdditions();
            form1.populateHrrates();
            form1.Updateboxes();
            //form1.BoxUpdate();
            //form1.Show();


        }

        private void RaddBtn_Click(object sender, EventArgs e)
        {
            string strinput = Hrnametb.Text;
            double Weight;
            bool isNumeric = double.TryParse(this.Hrratetb.Text, out Weight);

            Hrrates hrrateItem = new Hrrates();
            hrrateItem.hrsname = strinput;
            hrrateItem.hrsWeight = Weight;
            //newBe.lists.HrrateList.Add(hrrateItem);
            newBe.lists.AddtoHrrates(hrrateItem);
            populateHrrates();

            Hrratetb.Text = "";
            Hrnametb.Text = "";

            Main form1 = (Main)this.Owner;
            form1.populateHrrates();
        }

        private void populateHrrates()
        {
            newBe.lists.RebuildHrRates();
            this.Ratelistlv.Items.Clear();
            foreach (Hrrates hritem in newBe.lists.HrrateList)
            {
                if (hritem.hrsname == "PcGns")
                    pcEGnstb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "PcMp")
                    pcEMptb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "PcEdgeGns")
                    pcEdgeGnstb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "PcEdgeMp")
                    pcEdgeMptb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "GrnSealFullpr")
                    GnsFulltb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "GrnSealLightpr")
                    GnsLighttb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "GrnSealMedpr")
                    GnsMedtb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "MpSealLightpr")
                    MpLighttb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "MpSealMedpr")
                    MpMedtb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "MpSealFullpr")
                    MpFulltb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "Costmarkup")
                    Costmarkuptb.Text = hritem.hrsWeight.ToString();
                else if (hritem.hrsname == "HourMulti")
                    HoursMultitb.Text = hritem.hrsWeight.ToString();

                else
                {
                    string[] newArr = new string[12];
                    newArr[0] = hritem.hrsname;
                    newArr[1] = hritem.hrsWeight.ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Ratelistlv.Items.Add(itm);
                }
            }
        }
        private void populatAdditions()
        {
            newBe.lists.RebuildAdditions();
            this.Addlv.Items.Clear();
            foreach (string stritem in newBe.lists.AdditionsList)
            {
                this.Addlv.Items.Add(stritem);
            }
        }
        private void HrRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (Ratelistlv.SelectedIndices.Count > 0)
            {
                Selected = Ratelistlv.SelectedIndices[0];
                item = Ratelistlv.Items[Selected].Text;
                newBe.lists.DelItemFromHrrates(item);
            }
            populateHrrates();
        }

        private void Updatebtn_Click(object sender, EventArgs e)
        {
            foreach (Hrrates hritem in newBe.lists.HrrateList)
            {
                if (hritem.hrsname == "PcEdgeGns")
                {
                    hritem.hrsWeight = double.Parse(pcEMptb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "GrnSealLightpr")
                {
                    hritem.hrsWeight = double.Parse(GnsLighttb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "GrnSealMedpr")
                {
                    hritem.hrsWeight = double.Parse(GnsMedtb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "GrnSealFullpr")
                {
                    hritem.hrsWeight = double.Parse(GnsFulltb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "PcMp")
                {
                    hritem.hrsWeight = double.Parse(pcEMptb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "PcEdgeMp")
                {
                    hritem.hrsWeight = double.Parse(pcEdgeMptb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "MpSealLightpr")
                {
                    hritem.hrsWeight = double.Parse(MpLighttb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                } 
                if (hritem.hrsname == "PcGns")
                {
                    hritem.hrsWeight = double.Parse(pcEGnstb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "MpSealFullpr")
                {
                    hritem.hrsWeight = double.Parse(MpFulltb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "MpSealMedpr")
                {
                    hritem.hrsWeight = double.Parse(MpMedtb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
                if (hritem.hrsname == "HourMulti") 
                {
                    hritem.hrsWeight = double.Parse(HoursMultitb.Text);
                    newBe.lists.UpdateHrrates(hritem);                    
                }
                if (hritem.hrsname == "Costmarkup")
                {
                    hritem.hrsWeight = double.Parse(Costmarkuptb.Text);
                    newBe.lists.UpdateHrrates(hritem);
                }
            }
        }

        private void Addaddbtn_Click(object sender, EventArgs e)
        {
            if(Addnametb.Text != "")
            {
                //newBe.lists.AdditionsList.Add(Addnametb.Text);
                newBe.lists.UpdateAdditions(Addnametb.Text);
            }
            populatAdditions();
        }

        private void AddRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (Addlv.SelectedIndices.Count > 0)
            {
                Selected = Addlv.SelectedIndices[0];
                item = Addlv.Items[Selected].Text;
                newBe.lists.DelItemFromAdditions(item);
            }
            populatAdditions();
        }
    }
}
