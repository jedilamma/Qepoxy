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
    public partial class AIOne : Form
    {
        Backend newBe;
        //List<Priceitem> ItemList = new List<Priceitem>();
        //List<Pack> LocalpackList = new List<Pack>();
        List<Priceitem> TempProductList = new List<Priceitem>();
        List<Pack> BuildPacklist = new List<Pack>();
        //List<ProductSystem> LocalSystemSList = new List<ProductSystem>();

        //List<ProductSystem> ProductList = new List<ProductSystem>();
        List<Priceitem> LocalproductList = new List<Priceitem>();
        //List<ProductSystem> systemlist = new List<ProductSystem>();

        public AIOne(Backend Backend)
        {
            InitializeComponent();
            newBe = Backend;
            newBe.lists.Addtolog("All in one opened");
            this.PacksAvailLv.Scrollable = true;
            this.PacksAvailLv.View = View.Details;
            this.PacksAvailLv.GridLines = true;
            this.PacksAvailLv.FullRowSelect = true;
            this.PacksAvailLv.Columns.Add("Product", 160);
            this.PacksAvailLv.Columns.Add("Price", 55);
            this.PacksAvailLv.Columns.Add("Weight", 45);
            this.PacksAvailLv.Columns.Add("Volume", 45);
            this.PacksAvailLv.Columns.Add("Coverage", 45);
            this.PacksAvailLv.Columns.Add("Unit", 45);

            this.TempProdlv.Scrollable = true;
            this.TempProdlv.View = View.Details;
            this.TempProdlv.GridLines = true;
            this.TempProdlv.FullRowSelect = true;
            this.TempProdlv.Columns.Add("Product", 160);
            this.TempProdlv.Columns.Add("Price", 55);
            this.TempProdlv.Columns.Add("Weight", 45);
            this.TempProdlv.Columns.Add("Volume", 45);
            this.TempProdlv.Columns.Add("Coverage", 45);
            this.TempProdlv.Columns.Add("Unit", 45);

            this.Systemslv.Scrollable = true;
            this.Systemslv.View = View.Details;
            this.Systemslv.GridLines = true;
            this.Systemslv.FullRowSelect = true;
            this.Systemslv.Columns.Add("Name", 220);
            this.Systemslv.Columns.Add("#", 35);
            this.Systemslv.Columns.Add("Type", 75);

            this.PbuilderLv.Scrollable = true;
            this.PbuilderLv.View = View.Details;
            this.PbuilderLv.GridLines = true;
            this.PbuilderLv.FullRowSelect = true;
            this.PbuilderLv.Columns.Add("Name", 280);
            this.PbuilderLv.Columns.Add("#", 35);
            //this.PbuilderLv.Columns.Add("Options", 160);

            //this.Packtypecb.Items.AddRange(new object[] { "Floor" });
            this.Packtypecb.Items.AddRange(new object[] { "Floor-PUC" });
            this.Packtypecb.Items.AddRange(new object[] { "Floor-RC" });
            this.Packtypecb.Items.AddRange(new object[] { "Floor-BCF" });
            this.Packtypecb.Items.AddRange(new object[] { "Floor-ESL" });
            this.Packtypecb.Items.AddRange(new object[] { "Coving" });
            this.Packtypecb.Items.AddRange(new object[] { "LineMarking" });
            this.Packtypecb.Items.AddRange(new object[] { "Chevron" });
            this.Packtypecb.Items.AddRange(new object[] { "Subfill" });
            this.Packtypecb.Items.AddRange(new object[] { "Polished-Concrete" });
            this.Packtypecb.Items.AddRange(new object[] { "Preperation" });
            this.Packtypecb.Items.AddRange(new object[] { "Joints" });
            this.Packtypecb.Items.AddRange(new object[] { "Consumable" });
            //this.Packtypecb.Items.AddRange(new object[] { "Additions" });
            this.Packtypecb.Items.AddRange(new object[] { "Not-listed" });

            //LoadSystems();
            //LoadSystems();
            //newBe.lists.LoadSQLData();

            //ItemList = newBe.lists.PriceList;
            //LocalpackList = newBe.lists.PackList;
            //LocalSystemSList = newBe.lists.SystemsList;

            this.Productlv.Scrollable = true;
            this.Productlv.View = View.Details;
            this.Productlv.GridLines = true;
            this.Productlv.FullRowSelect = true;
            this.Productlv.Columns.Add("Product", 180);
            this.Productlv.Columns.Add("Price", 50);
            this.Productlv.Columns.Add("Weight", 50);
            this.Productlv.Columns.Add("Volume", 50);
            this.Productlv.Columns.Add("Coverage", 50);
            this.Productlv.Columns.Add("Unit", 35);
            this.Productlv.Columns.Add("ProductDate", 90);

            Pbltb.Enabled = false;
            Pbctb.Enabled = false;
            Pbcbtb.Enabled = false;
            Pcltb.Enabled = false;
            Pcctb.Enabled = false;
            Pccctb.Enabled = false;
            Pdltb.Enabled = false;
            Pdctb.Enabled = false;
            Pdcdtb.Enabled = false;

            Packb1tb.Enabled = false;
            Packc1tb.Enabled = false;
            Packd1tb.Enabled = false;

            Populate();

            this.FormClosing += AIOne_FormClosing;
        }
        private void AIOne_FormClosing(Object sender, FormClosingEventArgs e)
        {
            //UpdatePricesInLists();
            //newBe.lists.SavePriceItemsData();
            //System.Threading.Thread.Sleep(100);
            //newBe.lists.SavePackData();
            //System.Threading.Thread.Sleep(100);
            //newBe.lists.SaveSystemsData();
            //System.Threading.Thread.Sleep(100);

            //LocalSystemSList.Sort((x, y) => x.BaseMat.CompareTo(y.BaseMat));
            //LocalpackList.Sort((x, y) => x.Name.CompareTo(y.Name));
            //ItemList.Sort((x, y) => x.Name.CompareTo(y.Name));

            //newBe.lists.PriceList = ItemList;
            //newBe.lists.PackList = LocalpackList;
            //newBe.lists.SystemsList = LocalSystemSList;

            Main form1 = (Main)this.Owner;
            newBe.lists.RelinkPackLinks();
            newBe.lists.RelinkSystemDataLinks();
            form1.Populate();
            form1.UpdateFloorsysBox();
            form1.UpdatePrepsysBox();            
            form1.UpdateBCFloorBox();
            form1.UpdateESLloorBox();
            form1.UpdateRCFloorBox();
            form1.UpdateConsumesBox();
            //newBe.lists.Addtolog("All in one closed");

        }
        public void LoadSystems()
        {
            Populate();
        }
        private void ProdsaveBtn_Click(object sender, EventArgs e)
        {
            this.errorPR1.Clear();
            bool noError = true;
            double errOut;
            if (errorCheckName(ProdNametb.Text))
            {
                this.errorPR1.SetError(this.ProdNametb, "contains , bad charecter");
                noError = false;
            }
            bool errCheck1 = double.TryParse(ProdPrtb.Text, out errOut);
            if (!errCheck1)
            {
                this.errorPR1.SetError(this.ProdPrtb, "All Charecters should be numerical");
                noError = false;
            }
            if (errorCheckName(ProdPrtb.Text))
            {
                this.errorPR1.SetError(this.ProdPrtb, "contains , bad charecter");
                noError = false;
            }
            bool errCheck2 = double.TryParse(ProdWeighttb.Text, out errOut);
            if (!errCheck2)
            {
                this.errorPR1.SetError(this.ProdWeighttb, "All Charecters should be numerical");
                noError = false;
            }
            if (errorCheckName(ProdWeighttb.Text))
            {
                this.errorPR1.SetError(this.ProdWeighttb, "contains , bad charecter");
                noError = false;
            }
            bool errCheck3 = double.TryParse(Voltb.Text, out errOut);
            if (!errCheck3)
            {
                this.errorPR1.SetError(this.Voltb, "All Charecters should be numerical");
                noError = false;
            }
            if (errorCheckName(Voltb.Text))
            {
                this.errorPR1.SetError(this.Voltb, "contains , bad charecter");
                noError = false;
            }
            bool errCheck4 = double.TryParse(Covtb.Text, out errOut);
            if (!errCheck4)
            {
                this.errorPR1.SetError(this.Covtb, "All Charecters should be numerical");
                noError = false;
            }
            if (errorCheckName(Covtb.Text))
            {
                this.errorPR1.SetError(this.Covtb, "contains , bad charecter");
                noError = false;
            }
            if (errorCheckName(ProdUnittb.Text))
            {
                this.errorPR1.SetError(this.ProdUnittb, "contains , bad charecter");
                noError = false;
            }
            
            if (noError)
            { 
                if (!createProdItem())
                {
                    Populate();
                }
            }

            //UpdatePricesInLists();
        }

        //public void UpdatePricesInLists()
        //{
        //    foreach (ProductSystem psys in LocalSystemSList)
        //    {
        //        foreach (Pack pItem in psys.PackList)
        //        {
        //            foreach (Priceitem prItem in pItem.PriceItemList)
        //            {
        //                foreach (Priceitem localprItem in ItemList)
        //                {
        //                    if (prItem.Name == localprItem.Name)
        //                    {
        //                        prItem.Coverage = localprItem.Coverage;
        //                        prItem.Meterage = localprItem.Meterage;
        //                        prItem.Price = localprItem.Price;
        //                        prItem.Unit = localprItem.Unit;
        //                        prItem.Volume = localprItem.Volume;
        //                        prItem.Weight = localprItem.Weight;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    foreach (Pack localpackItem in LocalpackList)
        //    {
        //        foreach (Priceitem prItem in localpackItem.PriceItemList)
        //        {
        //            foreach (Priceitem localprItem in ItemList)
        //            {
        //                if (prItem.Name == localprItem.Name)
        //                {
        //                    prItem.Coverage = localprItem.Coverage;
        //                    prItem.Meterage = localprItem.Meterage;
        //                    prItem.Price = localprItem.Price;
        //                    prItem.Unit = localprItem.Unit;
        //                    prItem.Volume = localprItem.Volume;
        //                    prItem.Weight = localprItem.Weight;
        //                }
        //            }
        //        }
        //    }
        //}

        public bool createProdItem()
        {
            bool error = false;
            if (ProdNametb.Text != "")
            {
                Priceitem itm = new Priceitem();
                itm.Name = ProdNametb.Text;

                double price;
                bool PisNumeric = double.TryParse(this.ProdPrtb.Text, out price);
                if (PisNumeric)
                    itm.Price = price;
                else
                {
                    error = true;
                }
                if (ProdWeighttb.Text == "")
                {
                    itm.Weight = 0;
                }
                else
                {
                    double weight;
                    bool WisNumeric = double.TryParse(this.ProdWeighttb.Text, out weight);
                    if (WisNumeric)
                        itm.Weight = weight;
                    else
                        error = true;
                }
                if (Voltb.Text == "")
                {
                    itm.Volume = 0;
                }
                else
                {
                    double vol;
                    bool VisNumeric = double.TryParse(this.Voltb.Text, out vol);
                    if (VisNumeric)
                        itm.Volume = vol;
                    else
                        error = true;
                }

                if (Covtb.Text == "")
                {
                    itm.Coverage = 0;
                }
                else
                {
                    double cov;
                    bool CisNumeric = double.TryParse(this.Covtb.Text, out cov);
                    if (CisNumeric)
                        itm.Coverage = cov;
                    else
                        error = true;
                }
                if (ProdUnittb.Text != "")
                {
                    itm.Unit = ProdUnittb.Text;
                }
                itm.productDate = DateTime.Now.ToString("dd/MM/yyyy");
                //ItemList.Add(itm);
                newBe.lists.AdditemToPriceList(itm);
                int Packschanged = 0;
                int Systemschanged = 0;
                foreach (Pack pitem in newBe.lists.PackList)
                {
                    foreach(Priceitem  pritem in pitem.PriceItemList)
                    {
                        if (pritem.Name == itm.Name)
                        {
                            pritem.Price = itm.Price;
                            pritem.Weight = itm.Weight;
                            pritem.Volume = itm.Volume;
                            pritem.Unit = itm.Unit;
                            pritem.Coverage = itm.Coverage;
                            pritem.productDate = DateTime.Now.ToString("MM/dd/yyyy"); 
                            Packschanged++;
                        }
                    }
                }
                foreach (ProductSystem psys in newBe.lists.SystemsList)
                {
                    foreach (Pack pitem in psys.PackList)
                    {
                        foreach (Priceitem pritem in pitem.PriceItemList)
                        {
                            if (pritem.Name == itm.Name)
                            {
                                pritem.Price = itm.Price;
                                pritem.Weight = itm.Weight;
                                pritem.Volume = itm.Volume;
                                pritem.Unit = itm.Unit;
                                pritem.Coverage = itm.Coverage;
                                pritem.productDate = DateTime.Now.ToString("MM/dd/yyyy");
                                Systemschanged++;
                            }
                        }
                    }
                }
                // DialogResult dr;
                //dr = MessageBox.Show("Changing "+ ProdNametb.Text + " Has effected " + Packschanged + " and "+ Systemschanged + " Systems.","Pack Coverages Remain The Same",MessageBoxButtons.OK);
                Displaylbl.Text = "Changing " + ProdNametb.Text + " Has effected " + Packschanged + " and " + Systemschanged + " Systems.";
                resetlbl();

                ProdNametb.Text = "";
                Covtb.Text = "";
                Voltb.Text = "";
                ProdWeighttb.Text = "";
                ProdPrtb.Text = "";
                ProdUnittb.Text = "";

            }
            return error;
        }

        async private void resetlbl()
        {
            await Task.Delay(10000);
            Displaylbl.Text = "";
        }

        private bool createPackItem()
        {
            bool error = false;
            Pack newPackItem = new Pack();
            if (Packnametb.Text != "")
            {
                newPackItem.Name = Packnametb.Text;
                foreach (Priceitem item in TempProductList)
                {
                    newPackItem.AddChild(item);
                }

                double Pr;
                bool PrisNumeric = double.TryParse(Pricetb.Text, out Pr);
                if (PrisNumeric)
                    newPackItem.Price = Pr;
                else
                    error = true;

                double We;
                bool WeisNumeric = double.TryParse(weighttb.Text, out We);
                if (WeisNumeric)
                    newPackItem.Weight = We;
                else
                    error = true;

                double Vo;
                bool VoisNumeric = double.TryParse(volumetb.Text, out Vo);
                if (VoisNumeric)
                    newPackItem.Volume = Vo;
                else
                    error = true;

                double Cov;
                bool CovisNumeric = double.TryParse(Coveragetb.Text, out Cov);
                if (VoisNumeric)
                    newPackItem.Coverage = Cov;
                else
                    error = true;

                newPackItem.Unit = Unittb.Text;

                if (!error)
                {
                    newBe.lists.AdditemToPackList(newPackItem);
                    
                    foreach (ProductSystem psys in newBe.lists.SystemsList)
                    {
                        foreach (Pack pitem in psys.PackList)
                        {
                            if (pitem.Name == Packnametb.Text)
                            {
                                pitem.Price = Pr;
                                pitem.Weight = We;
                                pitem.Volume = Vo;                                
                                pitem.Coverage = Cov;
                                pitem.Unit = Unittb.Text;
                            }
                        }
                    }
                    PopulatePacksList();
                }

            }
            return error;
        }
        private void ProductEdit1_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Productlv.SelectedIndices.Count > 0)
            {
                Selected = Productlv.SelectedIndices[0];
                string item = Productlv.Items[Selected].Text;
                foreach (Priceitem PItem in newBe.lists.PriceList)
                {
                    if (PItem.Name == item)
                    {
                        ProdNametb.Text = item;
                        Covtb.Text = PItem.Coverage.ToString();
                        Voltb.Text = PItem.Volume.ToString();
                        ProdWeighttb.Text = PItem.Weight.ToString();
                        ProdPrtb.Text = PItem.Price.ToString();
                        ProdUnittb.Text = PItem.Unit;
                    }
                }
                //createItem();
                //Populate();
            }
        }
        

        private void CHcalcbtn_Click(object sender, EventArgs e)
        {
            if (CHAreatb.Text != "")
            {
                double Pa = 0;
                double Pb = 0;
                double Pc = 0;
                double Pd = 0;
                double Area = double.Parse(CHAreatb.Text);
                double Meterage = 0;
                if (ChMeteragetb.Text != "")
                    Meterage = double.Parse(ChMeteragetb.Text);
                if (Paltb.Text != "")
                    Pa = double.Parse(Paltb.Text);
                if (Pbltb.Text != "")
                    Pb = double.Parse(Pbltb.Text);
                if (Pcltb.Text != "")
                    Pc = double.Parse(Pcltb.Text);
                if (Pdltb.Text != "")
                    Pd = double.Parse(Pdltb.Text);

                if (Pbcb.Checked && Pccb.Checked && Pdcb.Checked)
                {
                    double TotalLitres = Pa + Pb + Pc;
                    double PaLitre = (Pa / TotalLitres);
                    double PbLitre = (Pb / TotalLitres);
                    double PdLitre = (Pd / TotalLitres);
                    double PcLitre = (Pc / TotalLitres);

                    double Packs = Meterage / (TotalLitres * ((PaLitre * (Area / TotalLitres)) + (PdLitre * (Area / TotalLitres)) + (PbLitre * (Area / TotalLitres)) + (PcLitre * (Area / TotalLitres))));

                    Packa2tb.Text = Math.Round((Packs), 2).ToString();

                    Packa1tb.Text = Math.Round((Packs * Pa), 6).ToString();
                    Packb1tb.Text = Math.Round((Packs * Pb), 6).ToString();
                    Packc1tb.Text = Math.Round((Packs * Pc), 6).ToString();
                    Packd1tb.Text = Math.Round((Packs * Pd), 6).ToString();

                    Pactb.Text = Math.Round((PaLitre * (Area / TotalLitres)), 6).ToString();
                    Pacatb.Text = Math.Round((PaLitre * (TotalLitres / Area)), 6).ToString();

                    Pbctb.Text = Math.Round((PbLitre * (Area / TotalLitres)), 6).ToString();
                    Pbcbtb.Text = Math.Round((PbLitre * (TotalLitres / Area)), 6).ToString();

                    Pcctb.Text = Math.Round((PcLitre * (Area / TotalLitres)), 6).ToString();
                    Pccctb.Text = Math.Round((PcLitre * (TotalLitres / Area)), 6).ToString();

                    Pdctb.Text = Math.Round((PdLitre * (Area / TotalLitres)), 6).ToString();
                    Pdcdtb.Text = Math.Round((PdLitre * (TotalLitres / Area)), 6).ToString();
                }
                else if (Pbcb.Checked && Pccb.Checked)
                {
                    double TotalLitres = Pa + Pb + Pc;
                    double PaLitre = (Pa / TotalLitres);
                    double PbLitre = (Pb / TotalLitres);
                    double PcLitre = (Pc / TotalLitres);

                    double Packs = Meterage / (TotalLitres * ((PaLitre * (Area / TotalLitres)) + (PbLitre * (Area / TotalLitres)) + (PcLitre * (Area / TotalLitres))));

                    Packa2tb.Text = Math.Round((Packs), 2).ToString();

                    Packa1tb.Text = Math.Round((Packs * Pa), 6).ToString();
                    Packb1tb.Text = Math.Round((Packs * Pb), 6).ToString();
                    Packc1tb.Text = Math.Round((Packs * Pc), 6).ToString();

                    Pactb.Text = Math.Round((PaLitre * (Area / TotalLitres)), 6).ToString();
                    Pacatb.Text = Math.Round((PaLitre * (TotalLitres / Area)), 6).ToString();

                    Pbctb.Text = Math.Round((PbLitre * (Area / TotalLitres)), 6).ToString();
                    Pbcbtb.Text = Math.Round((PbLitre * (TotalLitres / Area)), 6).ToString();

                    Pcctb.Text = Math.Round((PcLitre * (Area / TotalLitres)), 6).ToString();
                    Pccctb.Text = Math.Round((PcLitre * (TotalLitres / Area)), 6).ToString();

                }
                else if (Pbcb.Checked)
                {
                    double TotalLitres = Pa + Pb;
                    double PaLitre = (Pa / TotalLitres);
                    double PbLitre = (Pb / TotalLitres);

                    double Packs = Meterage / (TotalLitres * ((PaLitre * (Area / TotalLitres)) + (PbLitre * (Area / TotalLitres))));

                    Packa2tb.Text = Math.Round((Packs), 2).ToString();

                    Packa1tb.Text = Math.Round((Packs * Pa), 6).ToString();
                    Packb1tb.Text = Math.Round((Packs * Pb), 6).ToString();

                    Double CoverageA = (PaLitre * (Area / TotalLitres));
                    Pactb.Text = Math.Round(CoverageA, 6).ToString();
                    Pacatb.Text = Math.Round((PaLitre * (TotalLitres / Area)), 6).ToString();

                    Double CoverageB = (PbLitre * (Area / TotalLitres));
                    Pbctb.Text = Math.Round(CoverageB, 6).ToString();
                    Pbcbtb.Text = Math.Round((PbLitre * (TotalLitres / Area)), 6).ToString();

                }
                else
                {
                    //double pka = 0;

                    double TotalLitres = Pa;
                    double PaLitre = (Pa / TotalLitres);

                    double Packs = Meterage / (TotalLitres * ((PaLitre * (Area / TotalLitres))));

                    Packa2tb.Text = Math.Round((Packs), 6).ToString();

                    Packa1tb.Text = Math.Round((Packs * Pa), 6).ToString();

                    Pactb.Text = Math.Round((PaLitre * (Area / TotalLitres)), 6).ToString();
                    Pacatb.Text = Math.Round((PaLitre * (TotalLitres / Area)), 6).ToString();
                    //Meterage / pk
                    // pItem.Meterage / pk * prItem.Volume
                }
            }
        }

        public void PopulateSysBuilderList()
        {
            
            this.PbuilderLv.Items.Clear();
            BuildPacklist.Sort((x, y) => x.Name.CompareTo(y.Name));

            foreach (Pack pItem in BuildPacklist)
            {
                string[] newArr = new string[12];
                newArr[0] = pItem.Name;
                newArr[1] = pItem.PriceItemList.Count().ToString();                
                ListViewItem itm = new ListViewItem(newArr);
                this.PbuilderLv.Items.Add(itm);
            }
        }

        public void PopulateProdList()
        {
            newBe.lists.RebuildPriceData();
            this.Productlv.Items.Clear();            
            foreach (Priceitem prod in newBe.lists.PriceList)
            {
                string[] newArr = new string[12];
                newArr[0] = prod.Name;
                newArr[1] = prod.Price.ToString();
                newArr[2] = prod.Weight.ToString();
                newArr[3] = prod.Volume.ToString();
                newArr[4] = prod.Coverage.ToString();
                newArr[5] = prod.Unit;
                newArr[6] = prod.productDate;
                ListViewItem itm = new ListViewItem(newArr);
                this.Productlv.Items.Add(itm);
            }
        }
        public void PopulateTempProdList()
        {
            this.TempProdlv.Items.Clear();
            foreach (Priceitem Pritem in TempProductList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pritem.Name;
                newArr[1] = Pritem.Price.ToString();
                newArr[2] = Pritem.Weight.ToString();
                newArr[3] = Pritem.Volume.ToString();
                newArr[4] = Pritem.Coverage.ToString();
                newArr[5] = Pritem.Unit;
                ListViewItem Pritm = new ListViewItem(newArr);
                this.TempProdlv.Items.Add(Pritm);
                //this.ProductsLv.Items.Add(Pitem.Name);
            }
        }

        public void PopulatePacksList()
        {
            newBe.lists.RebuildPackData();
            populatePacks();
        }
        public void populateSystems()
        {
            this.Systemslv.Items.Clear();
            foreach (ProductSystem psItem in newBe.lists.SystemsList)
            {
                if (psItem.BaseMat.ToUpper().Contains(SystemsSearchtb.Text.ToUpper()))
                {
                    string[] newArr = new string[12];
                    newArr[0] = psItem.BaseMat;
                    newArr[1] = psItem.PackList.Count().ToString();
                    newArr[2] = psItem.Type;
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Systemslv.Items.Add(itm);
                }
                else if (SystemsSearchtb.Text == "")
                {

                    string[] newArr = new string[12];
                    newArr[0] = psItem.BaseMat;
                    newArr[1] = psItem.PackList.Count().ToString();
                    newArr[2] = psItem.Type;
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Systemslv.Items.Add(itm);
                }
            }
        }
        public void PopulateSystemsList()
        {
            newBe.lists.RebuildProductSystems();
            populateSystems();

        }
        public void Populate()
        {
            //this.ProductsLv.Items.Clear();            
            PopulateProdList();
            PopulateTempProdList();
            PopulatePacksList();
            PopulateSysBuilderList();
            PopulateSystemsList();
            populatePriceview();
            populatePacks();

        }
        private void ClearPackList()
        {
            TempProductList.Clear();
            Packnametb.Text = "";
            Pricetb.Text = "";
            weighttb.Text = "";
            volumetb.Text = "";
            Coveragetb.Text = "";
            Unittb.Text = "";
            PopulatePacksList();
            PopulateTempProdList();
            //Populate();
        }
        private void PackSavebtn_Click(object sender, EventArgs e)
        {
            bool noError = true;
            this.errorPR1.Clear();
            double errOut;

            bool errCheck1 = double.TryParse(Pricetb.Text, out errOut);
            if (!errCheck1)
            {
                this.errorPR1.SetError(this.Pricetb, "Price Should Be Numerical Values Only");
                noError = false;
            }

            bool errCheck2 = double.TryParse(weighttb.Text, out errOut);if (!errCheck2) { this.errorPR1.SetError(this.weighttb, "Weight Should Be Numerical Values Only"); noError = false;    }
            bool errCheck3 = double.TryParse(volumetb.Text, out errOut);if (!errCheck3){this.errorPR1.SetError(this.volumetb, "Volume Should Be Numerical Values Only");noError = false;}
            bool errCheck4 = double.TryParse(Coveragetb.Text, out errOut);if (!errCheck4){this.errorPR1.SetError(this.Coveragetb, "Coverage Should Be Numerical Values Only");noError = false;}
            
            if (errorCheckName(Pricetb.Text))   {this.errorPR1.SetError(this.Pricetb, "contains , bad charecter");noError = false;}
            if (errorCheckName(weighttb.Text))  {this.errorPR1.SetError(this.weighttb, "contains , bad charecter");noError = false;}
            if (errorCheckName(volumetb.Text))  {this.errorPR1.SetError(this.volumetb, "contains , bad charecter");noError = false;}
            if (errorCheckName(Coveragetb.Text)){this.errorPR1.SetError(this.Coveragetb, "contains , bad charecter");noError = false;}
            if (errorCheckName(Unittb.Text))    {this.errorPR1.SetError(this.Unittb, "contains , bad charecter");noError = false;}
            if (errorCheckName(Packnametb.Text)){this.errorPR1.SetError(this.Packnametb, "contains , bad charecter");noError = false;}

            if (Packnametb.Text == "")
            {
                this.errorPR1.SetError(this.Packnametb, "Name Cannot Be Empty For It To Save");
                noError = false;
            }
            if (noError)
            {
                createPackItem();
                ClearPackList();
            }
            
        }

        private void Addbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Productlv.SelectedIndices.Count > 0)
            {
                Selected = Productlv.SelectedIndices[0];
                string item = Productlv.Items[Selected].Text;
                foreach (Priceitem PItem in newBe.lists.PriceList)
                {
                    if (PItem.Name == item)
                    {
                        Priceitem Newitem = new Priceitem();
                        Newitem = PItem;
                        TempProductList.Add(Newitem);
                    }
                }
                //createItem();
                //Populate();
                PopulateTempProdList();
            }
        }

        private void Removebtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (TempProdlv.SelectedIndices.Count > 0)
            {
                Selected = TempProdlv.SelectedIndices[0];
                string item = TempProdlv.Items[Selected].Text;
                foreach (Priceitem PItem in TempProductList)
                {
                    if (PItem.Name == item)
                    {
                        TempProductList.Remove(PItem);
                        //ClearList();
                        //Populate();
                        PopulateTempProdList();
                        return;
                    }
                }
            }
        }

        private void Clearbtn_Click(object sender, EventArgs e)
        {
            TempProductList.Clear();
            Packnametb.Text = "";
            Pricetb.Text = "";
            weighttb.Text = "";
            volumetb.Text = "";
            Coveragetb.Text = "";
            Unittb.Text = "";
            //Populate();
            PopulateTempProdList();
        }

        private void CalcBtn_Click(object sender, EventArgs e)
        {
            double TotalPr = 0;
            double TotalWeight = 0;
            double TotalVolume = 0;
            double TotalCoverage = 0;
            //double PackCoverage = 0;
            foreach (Priceitem prItem in TempProductList)
            {
                TotalPr += prItem.Price;
                TotalWeight += prItem.Weight;
                TotalVolume += prItem.Volume;
                TotalCoverage += prItem.Coverage;
            }
            Pricetb.Text = TotalPr.ToString();
            weighttb.Text = TotalWeight.ToString();
            volumetb.Text = TotalVolume.ToString();
            Coveragetb.Text = TotalCoverage.ToString();
        }

        private void PackEditbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (PacksAvailLv.SelectedIndices.Count > 0)
            {
                Pack newPackItem = new Pack();
                Selected = PacksAvailLv.SelectedIndices[0];
                string item = PacksAvailLv.Items[Selected].Text;

                foreach (Pack paItem in newBe.lists.PackList)
                {
                    if (paItem.Name == item)
                    {
                        newPackItem = paItem;
                    }
                }

                //newBe.lists.

                TempProductList.Clear();
                foreach (Priceitem Pitem in newPackItem.PriceItemList)
                {
                    TempProductList.Add(Pitem);
                }

                Packnametb.Text = newPackItem.Name;
                Pricetb.Text = newPackItem.Price.ToString();
                weighttb.Text = newPackItem.Weight.ToString();
                volumetb.Text = newPackItem.Volume.ToString();
                Coveragetb.Text = newPackItem.Coverage.ToString();
                Unittb.Text = newPackItem.Unit;
                //Populate();

                PopulateTempProdList();
                PopulatePacksList();
                populatePacks();
            }
        }

        private void PackRemoveBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            List<string> effectedsystemslist = new List<string>();

            if (PacksAvailLv.SelectedIndices.Count > 0)
            {
                Pack newPackItem = new Pack();
                Selected = PacksAvailLv.SelectedIndices[0];
                item = PacksAvailLv.Items[Selected].Text;
               
            }

            if (item != "")
            {
                foreach (ProductSystem psItem in newBe.lists.SystemsList)
                {
                    foreach (Pack paitem in psItem.PackList)
                    {
                        if (paitem.Name == item)
                        {
                            effectedsystemslist.Add(psItem.BaseMat);
                        }
                    }
                }

                bool ynresult = false;
                DialogResult dr;
                if (effectedsystemslist.Count() > 0)
                {
                    dr = MessageBox.Show("Deleting " + item + " will delete " + effectedsystemslist.Count() + " system(s).",
                                         "Are you Sure?", MessageBoxButtons.YesNo);
                    switch (dr)
                    {
                        case DialogResult.Yes: { ynresult = true; break; }
                        case DialogResult.No: { break; }
                    }
                }

                if (ynresult || effectedsystemslist.Count() == 0)
                {
                    newBe.lists.DelItemFromPackList(item);
                    if (effectedsystemslist.Count() > 0)
                    {

                        foreach (string str in effectedsystemslist)
                        {
                            newBe.lists.DelItemFromSystemsList(str);
                        }
                    }
                }                
            }

            //Populate();
            PopulatePacksList();
        }
        private void CreateAddbtn_Click(object sender, EventArgs e)
        {            
            int Selected = 0;
            if (PacksAvailLv.SelectedIndices.Count > 0)
            {
                Selected = PacksAvailLv.SelectedIndices[0];
                string item = PacksAvailLv.Items[Selected].Text;
                
                foreach (Pack PItem in newBe.lists.PackList)
                {
                    if (PItem.Name == item)
                    {
                        Pack newPack = new Pack();
                        newPack = PItem;
                        BuildPacklist.Add(newPack);
                    }
                }
                //createItem();
                //Populate();
                PopulateSysBuilderList();
            }
        }

        private void CreateRemovebtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (PbuilderLv.SelectedIndices.Count > 0)
            {
                Selected = PbuilderLv.SelectedIndices[0];
                string itemName = PbuilderLv.Items[Selected].Text;
                newBe.lists.Addtolog("C - Packs removed " + itemName );
                foreach (Pack PrItem in BuildPacklist)
                {
                    if (PrItem.Name == itemName)
                    {
                        BuildPacklist.Remove(PrItem);
                        //Populate();
                        PopulateSysBuilderList();
                        break;
                    }
                }
            }
        }

        private void CreateCancelbtn_Click(object sender, EventArgs e)
        {
            BuildPacklist.Clear();
            SNametb.Text = "";
            Packtypecb.Text = "";
            SNametb.Text = "";
            Packtypecb.Text = "";
            //Populate();
            PopulateSysBuilderList();
        }
        private bool errorCheckName(string inputErr)
        {
            bool error = false;
            if(inputErr.Contains(","))
            {
                error = true;
            }
            return error;
        }
        private void CreateSavebtn_Click(object sender, EventArgs e)
        {            
            bool noError = true;            
            this.errorPR1.Clear();

            if (errorCheckName(SNametb.Text))
            {
                this.errorPR1.SetError(this.SNametb, "contains , bad charecter");
                noError = false;
            }

            if (SNametb.Text == "")
            {
                this.errorPR1.SetError(this.SNametb, "Cannot Save Blank Name of System");
                noError = false;
            }

            if (Packtypecb.Text == "")
            {
                this.errorPR1.SetError(this.Packtypecb, "System Needs a type");
                noError = false;
            }

            if (noError)
            {
                ProductSystem psys = new ProductSystem();
                psys.BaseMat = SNametb.Text;
                psys.Type = Packtypecb.Text;
                
                foreach (Pack pitem in BuildPacklist)
                {
                    psys.PackList.Add(pitem);
                }
                BuildPacklist.Clear();
                newBe.lists.AdditemToSystemsList(psys);

                BuildPacklist.Clear();
                SNametb.Text = "";
                Packtypecb.Text = "";
                //newBe.lists.SaveSystemsData();
                //Populate();
                PopulateSysBuilderList();
                PopulateSystemsList();
                
            }
        }

        private void EditBtn_Click(object sender, EventArgs e)
        {   
            int Selected = 0;
            BuildPacklist.Clear();
            if (Systemslv.SelectedIndices.Count > 0)
            {
                Selected = Systemslv.SelectedIndices[0];
                string itemName = Systemslv.Items[Selected].Text;
                foreach (ProductSystem PsItem in newBe.lists.SystemsList)
                {
                    if (PsItem.BaseMat == itemName)
                    {
                        foreach (Pack Pitem in PsItem.PackList)
                        {
                            BuildPacklist.Add(Pitem);
                        }
                        SNametb.Text = PsItem.BaseMat;
                        Packtypecb.Text = PsItem.Type;
                        //Populate();
                        PopulateSysBuilderList();
                        PopulateSystemsList();
                        break;
                    }
                }
            }
        }


        private void ProdSysRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Systemslv.SelectedIndices.Count > 0)
            {
                Selected = Systemslv.SelectedIndices[0];
                string itemName = Systemslv.Items[Selected].Text;                

                //newBe.lists.SystemsList.Remove(PsItem);
                newBe.lists.DelItemFromSystemsList(itemName);
                PopulateSystemsList();
            }
        }

        private void Pbcb_CheckedChanged(object sender, EventArgs e)
        {
            if (Pbcb.Checked)
            {
                Pbltb.Enabled = true;
                Pbctb.Enabled = true;
                Pbcbtb.Enabled = true;
                Packb1tb.Enabled = true;
            }
            else
            {
                Pbltb.Enabled = false;
                Pbctb.Enabled = false;
                Pbcbtb.Enabled = false;
                Packb1tb.Enabled = false;
            }
        }

        private void Pccb_CheckedChanged(object sender, EventArgs e)
        {
            if (Pccb.Checked)
            {
                Pcltb.Enabled = true;
                Pcctb.Enabled = true;
                Pccctb.Enabled = true;
                Packc1tb.Enabled = true;
            }
            else
            {
                Pcltb.Enabled = false;
                Pcctb.Enabled = false;
                Pccctb.Enabled = false;
                Packc1tb.Enabled = false;
            }
        }

        private void Pdcb_CheckedChanged(object sender, EventArgs e)
        {
            if (Pdcb.Checked)
            {
                Pdltb.Enabled = true;
                Pdctb.Enabled = true;
                Pdcdtb.Enabled = true;
                Packd1tb.Enabled = true;
            }
            else
            {
                Pdltb.Enabled = false;
                Pdctb.Enabled = false;
                Pdcdtb.Enabled = false;
                Packd1tb.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // newBe.lists.PriceList = ItemList;
            //newBe.lists.PackList = LocalpackList;
            //newBe.lists.SystemsList = LocalSystemSList;
            Main form1 = (Main)this.Owner;
            form1.Populate();            
            form1.UpdateFloorsysBox();
        }
        private void populatePriceview()
        {
            this.Productlv.Items.Clear();
            foreach (Priceitem prod in newBe.lists.PriceList)
            {
                string[] newArr = new string[12];
                if (prod.Name.ToUpper().Contains(PriceViewSearchtb.Text.ToUpper()))
                {
                    newArr[0] = prod.Name;
                    newArr[1] = prod.Price.ToString();
                    newArr[2] = prod.Weight.ToString();
                    newArr[3] = prod.Volume.ToString();
                    newArr[4] = prod.Coverage.ToString();
                    newArr[5] = prod.Unit;
                    newArr[6] = prod.productDate;
                    ListViewItem itm = new ListViewItem(newArr);
                    Productlv.Items.Add(itm);
                }
                else if (PriceViewSearchtb.Text == "")
                {
                    newArr[0] = prod.Name;
                    newArr[1] = prod.Price.ToString();
                    newArr[2] = prod.Weight.ToString();
                    newArr[3] = prod.Volume.ToString();
                    newArr[4] = prod.Coverage.ToString();
                    newArr[5] = prod.Unit;
                    newArr[6] = prod.productDate;
                    ListViewItem itm = new ListViewItem(newArr);
                    Productlv.Items.Add(itm);
                }
                //this.FloorType_Cb.Items.AddRange(new object[] { val.BaseMat });
            }
        }
        private void PriceViewSearchtb_TextChanged(object sender, EventArgs e)
        {
            populatePriceview();            
        }

        public void populatePacks()
        {
            this.PacksAvailLv.Items.Clear();
            foreach (Pack pacitem in newBe.lists.PackList)
            {
                string[] newArr = new string[12];
                if (pacitem.Name.ToUpper().Contains(PackSearchtb.Text.ToUpper()))
                {
                    newArr[0] = pacitem.Name;
                    newArr[1] = pacitem.Price.ToString();
                    newArr[2] = pacitem.Weight.ToString();
                    newArr[3] = pacitem.Volume.ToString();
                    newArr[4] = pacitem.Coverage.ToString();
                    newArr[5] = pacitem.Unit;
                    ListViewItem itm = new ListViewItem(newArr);
                    PacksAvailLv.Items.Add(itm);
                }
                else if (PackSearchtb.Text == "")
                {
                    newArr[0] = pacitem.Name;
                    newArr[1] = pacitem.Price.ToString();
                    newArr[2] = pacitem.Weight.ToString();
                    newArr[3] = pacitem.Volume.ToString();
                    newArr[4] = pacitem.Coverage.ToString();
                    newArr[5] = pacitem.Unit;
                    ListViewItem itm = new ListViewItem(newArr);
                    PacksAvailLv.Items.Add(itm);
                }
            }

        }
        private void PackSearchtb_TextChanged(object sender, EventArgs e)
        {
            populatePacks();
        }

        private void SystemsSearchtb_TextChanged(object sender, EventArgs e)
        {
            populateSystems();
        }

        private void Prdelbtn_Click(object sender, EventArgs e)
        {


            int Selected = 0;
            string item = "";
            List<string> effectedPacklist = new List<string>();
            List<string> effectedsystemslist = new List<string>();

            if (Productlv.SelectedIndices.Count > 0)
            {
                Selected = Productlv.SelectedIndices[0];
                item = Productlv.Items[Selected].Text;
            }
            if (item != "")
            {
                foreach (Pack paItem in newBe.lists.PackList)
                {
                    foreach (Priceitem prItem in paItem.PriceItemList)
                    {
                        if (item == prItem.Name)
                        {
                            effectedPacklist.Add(paItem.Name);
                        }
                    }
                }

                foreach (ProductSystem psItem in newBe.lists.SystemsList)
                {
                    foreach (Pack paitem in psItem.PackList)
                    {
                        foreach (string str in effectedPacklist)
                        {
                            if (paitem.Name == str)
                            {
                                effectedsystemslist.Add(psItem.BaseMat);
                            }
                        }
                    }
                }

                bool ynresult = false;

                DialogResult dr;
                if (effectedPacklist.Count() > 0)
                {
                    dr = MessageBox.Show("Deleting " + item + " will delete " + effectedPacklist.Count() + " pack(s), and " + effectedsystemslist.Count() + " system(s) aswell",
                                         "Are you Sure?", MessageBoxButtons.YesNo);
                    switch (dr)
                    {
                        case DialogResult.Yes: { ynresult = true; break; }
                        case DialogResult.No: { break; }
                    }
                }

                if (ynresult || effectedPacklist.Count() == 0)
                {
                    newBe.lists.DelItemFromPriceList(item);

                    if (effectedPacklist.Count() > 0)
                    {
                        foreach (string str in effectedPacklist)
                        {
                            newBe.lists.DelItemFromPackList(str);
                        }
                        foreach (string str in effectedsystemslist)
                        {
                            newBe.lists.DelItemFromSystemsList(str);
                        }
                    }
                }
            }
            PopulateProdList();
            if (effectedPacklist.Count() > 0)
                PopulatePacksList();
            if (effectedsystemslist.Count() > 0)
                PopulateSystemsList();

        }

        //private void SysPackSavebtn_Click(object sender, EventArgs e)
        //{

        //}

        //private void SysPackEditbtn_Click(object sender, EventArgs e)
        //{
        //    int Selected = 0;
        //    string item = "";
        //    string Sysname = SNametb.Text;
        //    if (PbuilderLv.SelectedIndices.Count > 0)
        //    {
        //        Pack newPackItem = new Pack();
        //        Selected = PbuilderLv.SelectedIndices[0];
        //        item = PbuilderLv.Items[Selected].Text;
        //    }

        //    TempProductList.Clear();
        //    foreach (ProductSystem sys in newBe.lists.SystemsList)
        //    {
        //        if(sys.BaseMat == Sysname)
        //        {
        //            foreach(Pack pitem in sys.PackList)
        //            {
        //                if(pitem.Name == item)
        //                {
        //                    foreach (Priceitem psitem in pitem.PriceItemList)
        //                    {
        //                        TempProductList.Add(psitem);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    PopulateTempProdList();
        //}
    }
}
