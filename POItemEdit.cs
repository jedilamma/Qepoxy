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
    public partial class POItemEdit : Form
    {
        double m_packsAmnt;
        int m_index;
        public POItemEdit(int index, string Desc, double Qty, double Cost,double packsAmnt,string Unit)
        {
            InitializeComponent();

            this.Unitcb.Items.AddRange(new object[] { "L" });
            this.Unitcb.Items.AddRange(new object[] { "Kg" });
            this.Unitcb.Items.AddRange(new object[] { "m" });
            this.Unitcb.Items.AddRange(new object[] { "Pack" });
            this.Unitcb.Items.AddRange(new object[] { "Bag" }); 

            m_index = index;
            Desctb.Text = Desc;
            Packamnttb.Text = packsAmnt.ToString();
            Unitcb.Text = Unit;
            Qtytb.Text = Qty.ToString();
            CostPtb.Text = Cost.ToString();
            m_packsAmnt = packsAmnt;
        }
        private void Calc_Click(object sender, EventArgs e)
        {
            Pricetb.Text = ((double.Parse(Qtytb.Text) / m_packsAmnt)* double.Parse(CostPtb.Text)).ToString();
        }

        private void Savebtn_Click(object sender, EventArgs e)
        {
            PurchaseOrder Po = (PurchaseOrder)this.Owner;
            double priceOut = double.Parse(Pricetb.Text);
            Po.SaveItemTolist(m_index, Desctb.Text,double.Parse(Qtytb.Text), double.Parse(CostPtb.Text), priceOut,double.Parse(Packamnttb.Text),Unitcb.Text);
            this.Close();
        }

        private void Cancelbtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
