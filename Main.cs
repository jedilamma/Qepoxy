using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
//using LiveCharts; //Core of the library
//using LiveCharts.Wpf; //The WPF controls
//using LiveCharts.WinForms; //the WinForm wrappers

namespace Qepoxy
{
    public partial class Main : Form
    {
        Backend newBackend;
        List<SummaryOutput> SummaryList1 = new List<SummaryOutput>();
        List<string> HoursList = new List<string>();
        List<Priceitem> costList = new List<Priceitem>();

        List<Priceitem> PcCostList = new List<Priceitem>();
        List<Pack> chevronList = new List<Pack>();
        List<Pack> linemarkList = new List<Pack>();
        List<Pack> subfillList = new List<Pack>();
        List<Pack> coveList = new List<Pack>();
        List<Pack> PUCfloorList = new List<Pack>();
        List<Pack> RCfloorList = new List<Pack>();
        List<Pack> BCFfloorList = new List<Pack>();
        List<Pack> ESLfloorList = new List<Pack>();
        List<Pack> PCList = new List<Pack>();
        List<Pack> PrepList = new List<Pack>();
        List<Pack> Jointlist = new List<Pack>();
        List<Pack> ConsumesList = new List<Pack>();
        //List<Adds> AdditionalList = new List<Adds>();

        List<stagedhrs> stagehrslist = new List<stagedhrs>();
        List<string> TestSummaryList = new List<string>();
        List<Pack> TestListFloor = new List<Pack>();
        public List<Costing> CostingList = new List<Costing>();
        public List<Costing> PcCostingList = new List<Costing>();
        public List<CSstage> StageList = new List<CSstage>();
        public Main(Backend backend)
        {
            InitializeComponent();
            newBackend = backend;
            newBackend.lists.LoadSQLData();

            this.ProdDetailsLV.Scrollable = true;
            this.ProdDetailsLV.View = View.Details;
            this.ProdDetailsLV.GridLines = true;
            this.ProdDetailsLV.FullRowSelect = true;
            //this.ProdDetailsLV.Columns.Add("Kits", 120);
            this.ProdDetailsLV.Columns.Add("#", 70);
            this.ProdDetailsLV.Columns.Add("Units", 50);
            this.ProdDetailsLV.Columns.Add("Product", 240);
            this.ProdDetailsLV.Columns.Add("$ per Kit", 70);
            this.ProdDetailsLV.Columns.Add("Est Price", 100);

            this.Stagelv.Scrollable = true;
            this.Stagelv.View = View.Details;
            this.Stagelv.GridLines = true;
            this.Stagelv.FullRowSelect = true;
            this.Stagelv.Columns.Add("Stage", 80);
            this.Stagelv.Columns.Add("System Type", 110);
            this.Stagelv.Columns.Add("#Area", 50);
            this.Stagelv.Columns.Add("#Lm", 50);
            this.Stagelv.Columns.Add("Product", 200);
            this.Stagelv.Columns.Add("#", 100);

            this.PolishedConclv.Scrollable = true;
            this.PolishedConclv.View = View.Details;
            this.PolishedConclv.GridLines = true;
            this.PolishedConclv.FullRowSelect = true;
            this.PolishedConclv.Columns.Add("Product", 200);
            this.PolishedConclv.Columns.Add("Packs", 80);
            this.PolishedConclv.Columns.Add("Area", 80);
            this.PolishedConclv.Columns.Add("Edge L", 80);
            this.PolishedConclv.Columns.Add("Coverage", 50);

            this.FloorLv.Scrollable = true;
            this.FloorLv.View = View.Details;
            this.FloorLv.GridLines = true;
            this.FloorLv.FullRowSelect = true;
            this.FloorLv.Columns.Add("Product", 300);
            this.FloorLv.Columns.Add("Packs", 60);
            this.FloorLv.Columns.Add("Area", 60);
            this.FloorLv.Columns.Add("Coverage", 60);

            this.Hoursweightlv.Scrollable = true;
            this.Hoursweightlv.View = View.Details;
            this.Hoursweightlv.GridLines = true;
            this.Hoursweightlv.FullRowSelect = true;
            this.Hoursweightlv.Columns.Add("Method", 115);
            this.Hoursweightlv.Columns.Add("Rate", 50);

            this.SummaryLv.Scrollable = true;
            this.SummaryLv.View = View.Details;
            this.SummaryLv.GridLines = false;
            this.SummaryLv.FullRowSelect = true;
            this.SummaryLv.Columns.Add("", 180);
            this.SummaryLv.Columns.Add("", 180);
            this.SummaryLv.Columns.Add("", 130);

            this.HrCombolv.Scrollable = true;
            this.HrCombolv.View = View.Details;
            this.HrCombolv.GridLines = true;
            this.HrCombolv.FullRowSelect = true;
            this.HrCombolv.Columns.Add("Stage", 100);
            this.HrCombolv.Columns.Add("SysName", 150);
            this.HrCombolv.Columns.Add("Method", 150);
            this.HrCombolv.Columns.Add("Meterage", 80);
            this.HrCombolv.Columns.Add("Rate", 80);
            this.HrCombolv.Columns.Add("Hours Result", 100);

            this.EstHrsDgv.Columns.Add("Area", "m2");
            this.EstHrsDgv.Columns.Add("Length", "Lm");
            this.EstHrsDgv.Columns.Add("Stage", "Stage Name");
            this.EstHrsDgv.Columns.Add("SysName", "System Name");
            this.EstHrsDgv.Columns.Add("Method", "Method");
            this.EstHrsDgv.Columns.Add("Techs", "Tech Hrs");
            this.EstHrsDgv.Columns.Add("Techs", "Hours");
            this.EstHrsDgv.Columns[0].Width = 80;
            this.EstHrsDgv.Columns[1].Width = 80;
            this.EstHrsDgv.Columns[2].Width = 100;
            this.EstHrsDgv.Columns[3].Width = 275;

            this.RCLv.Scrollable = true;
            this.RCLv.View = View.Details;
            this.RCLv.GridLines = true;
            this.RCLv.FullRowSelect = true;
            this.RCLv.Columns.Add("Product", 220);
            this.RCLv.Columns.Add("Packs", 60);
            this.RCLv.Columns.Add("Area", 60);
            this.RCLv.Columns.Add("Coverage", 60);

            this.BCFLv.Scrollable = true;
            this.BCFLv.View = View.Details;
            this.BCFLv.GridLines = true;
            this.BCFLv.FullRowSelect = true;
            this.BCFLv.Columns.Add("Product", 220);
            this.BCFLv.Columns.Add("Packs", 60);
            this.BCFLv.Columns.Add("Area", 60);
            this.BCFLv.Columns.Add("Coverage", 60);

            this.Conslv.Scrollable = true;
            this.Conslv.View = View.Details;
            this.Conslv.GridLines = true;
            this.Conslv.FullRowSelect = true;
            this.Conslv.Columns.Add("Product", 220);
            this.Conslv.Columns.Add("Packs", 60);
            this.Conslv.Columns.Add("Area", 60);
            this.Conslv.Columns.Add("Coverage", 60);

            this.ESLLv.Scrollable = true;
            this.ESLLv.View = View.Details;
            this.ESLLv.GridLines = true;
            this.ESLLv.FullRowSelect = true;
            this.ESLLv.Columns.Add("Product", 220);
            this.ESLLv.Columns.Add("Packs", 60);
            this.ESLLv.Columns.Add("Area", 60);
            this.ESLLv.Columns.Add("Coverage", 60);

            this.Additionslv.Scrollable = true;
            this.Additionslv.View = View.Details;
            this.Additionslv.GridLines = true;
            this.Additionslv.FullRowSelect = true;
            this.Additionslv.Columns.Add("ItemName", 256);

            this.PrepListlv.Scrollable = true;
            this.PrepListlv.View = View.Details;
            this.PrepListlv.GridLines = true;
            this.PrepListlv.FullRowSelect = true;
            this.PrepListlv.Columns.Add("Product", 220);
            this.PrepListlv.Columns.Add("Packs", 60);
            this.PrepListlv.Columns.Add("Area", 60);
            this.PrepListlv.Columns.Add("Coverage", 60);

            this.Consumelv.Scrollable = true;
            this.Consumelv.View = View.Details;
            this.Consumelv.GridLines = true;
            this.Consumelv.FullRowSelect = true;
            this.Consumelv.Columns.Add("Product", 220);
            this.Consumelv.Columns.Add("Packs", 60);
            this.Consumelv.Columns.Add("#", 60);
            this.Consumelv.Columns.Add("Coverage", 60);

            this.Prepsyslv.Scrollable = true;
            this.Prepsyslv.View = View.Details;
            this.Prepsyslv.GridLines = true;
            this.Prepsyslv.FullRowSelect = true;
            this.Prepsyslv.Columns.Add("Method", 200);
            //this.Prepsyslv.Columns.Add("Packs", 30);
            //this.Prepsyslv.Columns.Add("Area", 30);
            //this.Prepsyslv.Columns.Add("Coverage", 30);

            this.Covelv.Scrollable = true;
            this.Covelv.View = View.Details;
            this.Covelv.GridLines = true;
            this.Covelv.FullRowSelect = true;
            this.Covelv.Columns.Add("Product", 145);
            this.Covelv.Columns.Add("Lm", 55);
            this.Covelv.Columns.Add("Price", 55);
            this.Covelv.Columns.Add("Coverage", 55);
            this.Covelv.Columns.Add("Units", 55);
            this.Covelv.Columns.Add("Type", 55);

            this.LineMarklv.Scrollable = true;
            this.LineMarklv.View = View.Details;
            this.LineMarklv.GridLines = true;
            this.LineMarklv.FullRowSelect = true;
            this.LineMarklv.Columns.Add("Product", 300);
            this.LineMarklv.Columns.Add("Length", 45);
            //this.LineMarklv.Columns.Add("Colour", 55);

            this.JointsLv.Scrollable = true;
            this.JointsLv.View = View.Details;
            this.JointsLv.GridLines = true;
            this.JointsLv.FullRowSelect = true;
            this.JointsLv.Columns.Add("Product", 300);
            this.JointsLv.Columns.Add("Length", 45);
            //this.LineMarklv.Columns.Add("Colour", 55);

            this.ChevronLv.Scrollable = true;
            this.ChevronLv.View = View.Details;
            this.ChevronLv.GridLines = true;
            this.ChevronLv.FullRowSelect = true;
            this.ChevronLv.Columns.Add("Product", 155);
            this.ChevronLv.Columns.Add("Area", 45);
            this.ChevronLv.Columns.Add("Price", 55);
            this.ChevronLv.Columns.Add("Units", 45);
            this.ChevronLv.Columns.Add("Colour", 55);

            this.SubfillLv.Scrollable = true;
            this.SubfillLv.View = View.Details;
            this.SubfillLv.GridLines = true;
            this.SubfillLv.FullRowSelect = true;
            this.SubfillLv.Columns.Add("Product", 220);
            this.SubfillLv.Columns.Add("Area", 50);
            this.SubfillLv.Columns.Add("Volume", 50);
            this.SubfillLv.Columns.Add("Unit", 50);
            this.SubfillLv.Columns.Add("Coverage", 50);

            this.Addslv.Scrollable = true;
            this.Addslv.View = View.Details;
            this.Addslv.GridLines = true;
            this.Addslv.FullRowSelect = true;
            this.Addslv.Columns.Add("ItemName", 120);
            this.Addslv.Columns.Add("Description", 170);
            this.Addslv.Columns.Add("Ammnt", 100);
            this.Addslv.Columns.Add("Cost", 50);

            this.FloorSysLv.Scrollable = true;
            this.FloorSysLv.View = View.Details;
            this.FloorSysLv.GridLines = true;
            this.FloorSysLv.FullRowSelect = true;
            this.FloorSysLv.Columns.Add("System", 220);
            this.FloorSysLv.Columns.Add("#P", 30);
            this.FloorSysLv.HideSelection = false;

            this.FloorRCLv.Scrollable = true;
            this.FloorRCLv.View = View.Details;
            this.FloorRCLv.GridLines = true;
            this.FloorRCLv.FullRowSelect = true;
            this.FloorRCLv.Columns.Add("System", 220);
            this.FloorRCLv.Columns.Add("#P", 30);
            this.FloorRCLv.HideSelection = false;

            this.FloorBCFLv.Scrollable = true;
            this.FloorBCFLv.View = View.Details;
            this.FloorBCFLv.GridLines = true;
            this.FloorBCFLv.FullRowSelect = true;
            this.FloorBCFLv.Columns.Add("System", 220);
            this.FloorBCFLv.Columns.Add("#P", 30);
            this.FloorBCFLv.HideSelection = false;

            this.FloorESLLv.Scrollable = true;
            this.FloorESLLv.View = View.Details;
            this.FloorESLLv.GridLines = true;
            this.FloorESLLv.FullRowSelect = true;
            this.FloorESLLv.Columns.Add("System", 220);
            this.FloorESLLv.Columns.Add("#P", 30);
            this.FloorESLLv.HideSelection = false;

            Hrattachbtn.FlatStyle = FlatStyle.Flat;
            Hrattachbtn.FlatAppearance.BorderColor = Color.Green;
            Hrattachbtn.FlatAppearance.BorderSize = 2;

            StageHrsclearbtn.FlatStyle = FlatStyle.Flat;
            StageHrsclearbtn.FlatAppearance.BorderColor = Color.Red;
            StageHrsclearbtn.FlatAppearance.BorderSize = 2;

            this.PCvarcb.Items.AddRange(new object[] { "Grind & Seal" });
            this.PCvarcb.Items.AddRange(new object[] { "Mechanical Polish" });

            this.Exposurecb.Items.AddRange(new object[] { "Light" });
            this.Exposurecb.Items.AddRange(new object[] { "Medium" });
            this.Exposurecb.Items.AddRange(new object[] { "Full" });

            //this.Chevtb.Enabled = false;

            this.ChvColourTb.Items.AddRange(new object[] { "White" });
            this.ChvColourTb.Items.AddRange(new object[] { "Gold" });
            this.ChvColourTb.Items.AddRange(new object[] { "Red" });
            this.ChvColourTb.Items.AddRange(new object[] { "Blue" });

            this.Quoteslv.Scrollable = true;
            this.Quoteslv.View = View.Details;
            this.Quoteslv.GridLines = true;
            this.Quoteslv.FullRowSelect = true;
            this.Quoteslv.Columns.Add("QuoteName", 324);
            this.Quoteslv.MouseDoubleClick += new MouseEventHandler(Quoteslv_MouseDoubleClick);

            this.StageHourslv.Scrollable = true;
            this.StageHourslv.View = View.Details;
            this.StageHourslv.GridLines = true;
            this.StageHourslv.FullRowSelect = true;
            this.StageHourslv.Columns.Add("Stage", 150);
            this.StageHourslv.Columns.Add("System Type", 200);
            this.StageHourslv.Columns.Add("#Qty Area", 70);
            this.StageHourslv.Columns.Add("#Qty Lm", 70);
            this.StageHourslv.Columns.Add("Product", 300);

            ConsAreatb.Text = "";
            ConsAreatb.Enabled = false;
            ConsMultitb.Enabled = true;
            ConsMultitb.Text = "1";

            Populate();
            populateQuotelists();
            UpdateFloorsysBox();
            UpdatePrepsysBox();
            UpdateBCFloorBox();
            UpdateESLloorBox();
            UpdateRCFloorBox();
            UpdateConsumesBox();
            populateHrrates();
            populatAdditions();

            this.FormClosing += Form1_FormClosing;
            this.Additiontb.Text = "0";
            //this.CostMarkuptb.Text = newBackend.lists.HrRates.CostMarkUp.ToString();
            //this.MainHrRatetb.Text = newBackend.lists.HrRates.HrMulti.ToString();
            totalCosttb.Text = "0";
            Esthrstotaltb.Text = "0";

            //PoliPassestb.Text = "1";
            ChvPritb.Enabled = false;
            ChvPritb.Text = "10";
            CostAddstb.Text = "0";
            finalCosttb.Text = "0";
            HoursCosttb.Text = "0";
            MatTotaltb.Text = "0";
            TotalHrsTb.Text = "0";
            PcTotaltb.Text = "0";

            ConsAreatb.Text = "";
            ConsMultitb.Enabled = true;
            ConsMultitb.Text = "1";

            Updateboxes();
            CombRescb.Checked = true;
            Pareacb.Checked = true;
        }
        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            //StartForm CF = (StartForm)this.Owner;
            //CF.CloseRoutine();
        }

        private void AddToCosting(double Packvol, double packsAmnt, double Ammount, string symbol, string inputname, double cost, double price)
        {
            bool wasinList = false;
            if (CombRescb.Checked == true)
            {
                foreach (Costing Costitem in CostingList)
                {
                    if (Costitem.inputname == inputname)
                    {
                        wasinList = true;
                        Costitem.packsAmnt += packsAmnt;
                        Costitem.Ammount += Ammount;
                        Costitem.symbol = symbol;
                        Costitem.cost = cost;
                        Costitem.Packvol = Packvol;
                        Costitem.price += price;
                    }
                }
                if (!wasinList)
                {
                    Costing nc = new Costing();
                    nc.packsAmnt = packsAmnt;
                    nc.Ammount = Ammount;
                    nc.symbol = symbol;
                    nc.inputname = inputname;
                    nc.cost = cost;
                    nc.Packvol = Packvol;
                    nc.price = price;
                    CostingList.Add(nc);
                }
            }
            else
            {
                Costing nc = new Costing();
                nc.packsAmnt = packsAmnt;
                nc.Ammount = Ammount;
                nc.symbol = symbol;
                nc.inputname = inputname;
                nc.cost = cost;
                nc.Packvol = Packvol;
                nc.price = price;
                CostingList.Add(nc);
            }
        }
        private void AddToStage(string stagestring, string Systemtype, double Area, double Lm, double Packvol, double packsAmnt, double Ammount, string symbol, string inputname, double cost, double price, string Packsysname)
        {
            //string stagestring1 = fsStagecb.Text;
            //string stagestring2 = CovStagecb.Text;
            bool notinbox = false;
            //bool notinbox2 = false;
            //foreach (var cbitem in fsStagecb.Items)
            if (PUCStagecb.Items.Contains(stagestring))
            {
                notinbox = true;
            }
            if (!notinbox)
            {
                this.PUCStagecb.Items.AddRange(new object[] { stagestring });
                this.CovStagecb.Items.AddRange(new object[] { stagestring });
                this.sfStagecb.Items.AddRange(new object[] { stagestring });
                this.lmStagecb.Items.AddRange(new object[] { stagestring });
                this.chvStagecb.Items.AddRange(new object[] { stagestring });
                this.pcStagecb.Items.AddRange(new object[] { stagestring });
                this.PrepStagecb.Items.AddRange(new object[] { stagestring });
                this.RCStagecb.Items.AddRange(new object[] { stagestring });
                this.BCFStagecb.Items.AddRange(new object[] { stagestring });
                this.ESLStagecb.Items.AddRange(new object[] { stagestring });
                this.JointStagecb.Items.AddRange(new object[] { stagestring });
                this.ConsAtoScb.Items.AddRange(new object[] { stagestring });
            }

            bool wasinList = false;
            //foreach (CSstage Stageditem in StageList)
            //{
            //    if (stagestring == Stageditem.Staging && Stageditem.inputname == inputname)
            //    {
            //        if (Stageditem.Systems == Systemtype)
            //        {
            //            if (Stageditem.Systems == "Consumable")
            //            {
            //                if (Stageditem.PackSystemName == Packsysname)
            //                {
            //                    wasinList = true;
            //                    Stageditem.Systems = Systemtype;
            //                    Stageditem.Staging = stagestring;
            //                    Stageditem.Area = Area;
            //                    Stageditem.Lm = Lm;
            //                    Stageditem.packsAmnt = packsAmnt;
            //                    Stageditem.Ammount = Ammount;
            //                    Stageditem.symbol = symbol;
            //                    Stageditem.cost = cost;
            //                    Stageditem.Packvol = Packvol;
            //                    Stageditem.price = price;
            //                    Stageditem.PackSystemName = Packsysname;
            //                }
            //            }
            //            if (Stageditem.Area == Area && Stageditem.Systems != "Consumable")
            //            {
            //                if (Stageditem.PackSystemName == Packsysname)
            //                {
            //                    wasinList = true;
            //                    Stageditem.Systems = Systemtype;
            //                    Stageditem.Staging = stagestring;
            //                    Stageditem.Area = Area;
            //                    Stageditem.Lm = Lm;
            //                    Stageditem.packsAmnt = packsAmnt;
            //                    Stageditem.Ammount += Ammount;
            //                    Stageditem.symbol = symbol;
            //                    Stageditem.cost = cost;
            //                    Stageditem.Packvol = Packvol;
            //                    Stageditem.price = price;
            //                    Stageditem.PackSystemName = Packsysname;
            //                }
            //            }
            //        }
            //    }
            //}
            if (!wasinList)
            {               
                CSstage CSst = new CSstage();
                CSst.Systems = Systemtype;
                CSst.Staging = stagestring;
                CSst.Area = Area;
                CSst.Lm = Lm;
                CSst.packsAmnt = packsAmnt;
                CSst.Ammount = Ammount;
                CSst.symbol = symbol;
                CSst.inputname = inputname;
                CSst.cost = cost;
                CSst.Packvol = Packvol;
                CSst.price = price;
                CSst.PackSystemName = Packsysname;
                StageList.Add(CSst);                            
            }
            importToHrsList();
        }
        private void AddToPcCosting(double Packvol, double packsAmnt, double Ammount, string symbol, string inputname, double cost, double price)
        {
            bool wasinList = false;
            if (CombRescb.Checked == true)
            {
                foreach (Costing Costitem in PcCostingList)
                {
                    if (Costitem.inputname == inputname)
                    {
                        wasinList = true;
                        Costitem.packsAmnt += packsAmnt;
                        Costitem.Ammount += Ammount;
                        Costitem.symbol = symbol;
                        Costitem.cost = cost;
                        Costitem.Packvol = Packvol;
                        Costitem.price += price;
                    }
                }
                if (!wasinList)
                {
                    Costing nc = new Costing();
                    nc.packsAmnt = packsAmnt;
                    nc.Ammount = Ammount;
                    nc.symbol = symbol;
                    nc.inputname = inputname;
                    nc.cost = cost;
                    nc.Packvol = Packvol;
                    nc.price = price;
                    PcCostingList.Add(nc);
                }
            }
            else
            {
                Costing nc = new Costing();
                nc.packsAmnt = packsAmnt;
                nc.Ammount = Ammount;
                nc.symbol = symbol;
                nc.inputname = inputname;
                nc.cost = cost;
                nc.Packvol = Packvol;
                nc.price = price;
                PcCostingList.Add(nc);
            }
        }
        public void populateHrrates()
        {
            newBackend.lists.RebuildHrRates();
            this.Hoursweightlv.Items.Clear();
            foreach (Hrrates hritem in newBackend.lists.HrrateList)
            {
                if (hritem.hrsname == "PcGns") { }
                else if (hritem.hrsname == "PcMp") { }
                else if (hritem.hrsname == "PcEdgeGns") { }
                else if (hritem.hrsname == "PcEdgeMp") { }
                else if (hritem.hrsname == "GrnSealFullpr") { }
                else if (hritem.hrsname == "GrnSealLightpr") { }
                else if (hritem.hrsname == "GrnSealMedpr") { }
                else if (hritem.hrsname == "MpSealFullpr") { }
                else if (hritem.hrsname == "MpSealLightpr") { }
                else if (hritem.hrsname == "MpSealMedpr") { }
                else if (hritem.hrsname == "Costmarkup") { }
                else if (hritem.hrsname == "HourMulti") { }
                else
                {
                    string[] newArr = new string[12];
                    newArr[0] = hritem.hrsname;
                    newArr[1] = hritem.hrsWeight.ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Hoursweightlv.Items.Add(itm);
                }
            }
        }
        public void updateHRScombo()
        {
            HrCombolv.Items.Clear();
            foreach (HrCombo hrcmbitem in newBackend.lists.HrComboList)
            {
                string[] newArr = new string[12];
                newArr[0] = hrcmbitem.Stage;
                newArr[1] = hrcmbitem.SysName;
                newArr[2] = hrcmbitem.Method;
                newArr[3] = hrcmbitem.Area.ToString();
                newArr[4] = hrcmbitem.Length.ToString();
                newArr[5] = hrcmbitem.Rate.ToString();
                newArr[6] = hrcmbitem.Result.ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.HrCombolv.Items.Add(itm);
            }
        }
        public void updateProdDetails()
        {
            ProdDetailsLV.Items.Clear();
            foreach (Costing Costitem in CostingList)
            {
                string[] newArr = new string[12];
                double cstamnt = Math.Round(Costitem.packsAmnt, 1);
                //newArr[0] = cstamnt.ToString() + " kits";
                newArr[0] = Costitem.Ammount.ToString();
                newArr[1] = Costitem.symbol;
                newArr[2] = Costitem.inputname;
                newArr[3] = "$" + Math.Round(Costitem.cost, 2).ToString();
                newArr[4] = "$" + Math.Round(Costitem.price, 2).ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.ProdDetailsLV.Items.Add(itm);
            }

            foreach (Costing Costitem in PcCostingList)
            {
                string[] newArr = new string[12];
                double cstamnt = Math.Round(Costitem.packsAmnt, 1);
                newArr[0] = cstamnt.ToString() + " kits";
                newArr[1] = Costitem.Ammount.ToString();
                newArr[2] = Costitem.symbol;
                newArr[3] = Costitem.inputname;
                newArr[4] = "$" + Math.Round(Costitem.cost, 2).ToString();
                newArr[5] = "$" + Math.Round(Costitem.price, 2).ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.ProdDetailsLV.Items.Add(itm);
            }
        }
        //private void AddtoSummary(string input)
        //{
        //    SummaryList.Add(input);
        //}
        public void showEstTotal()
        {
            double totalcost = new double();
            totalcost = 0;
            double hoursVal = new double();
            hoursVal = 0;
            foreach (Costing Costitem in CostingList)
            {
                if (Costitem.symbol == "Hour")
                {
                    Costitem.packsAmnt = Costitem.packsAmnt;
                    Costitem.Ammount = Costitem.Ammount;
                    Costitem.symbol = Costitem.symbol;
                    Costitem.cost = Costitem.cost;
                    hoursVal += Costitem.packsAmnt * Costitem.cost;
                }
                else
                {
                    Costitem.packsAmnt = Costitem.packsAmnt;
                    Costitem.Ammount = Costitem.Ammount;
                    Costitem.symbol = Costitem.symbol;
                    Costitem.cost = Costitem.cost;
                    totalcost += Costitem.packsAmnt * Costitem.cost;
                }
            }

            updateProdDetails();
            newBackend.lists.POCostList = CostingList;
            double Markup = 0;
            double hrrate = 0;
            if (this.MainHrRatetb.Text != "")
            {
                hrrate = double.Parse(this.MainHrRatetb.Text);
            }

            if (this.CostMarkuptb.Text != "")
            {
                Markup = (double.Parse(this.CostMarkuptb.Text) + 100) / 100;
            }


            double matTotal = totalcost;
            MatTotaltb.Text = matTotal.ToString();
            HoursCosttb.Text = hoursVal.ToString();
            //double CostAdds = double.Parse(CostAddstb.Text);
            double CostAdds = 0;
            bool isNumeric = double.TryParse(this.CostAddstb.Text, out CostAdds);
            if (!isNumeric)
            {
                this.CostAddstb.Text = "0";
            }
            //double newtotalcost = totalcost * Markup + (double.Parse(TotalHrsTb.Text) + double.Parse(Additiontb.Text)) * hrrate;
            double newtotalcost = (hoursVal + totalcost) * Markup + double.Parse(Additiontb.Text);
            //double Pctotalcost = 0;
            //foreach (Costing Costitem in PcCostingList)
            //{
            //    Costitem.packsAmnt = Math.Ceiling(Costitem.packsAmnt);
            //    Costitem.Ammount = Costitem.Ammount;
            //    Costitem.symbol = Costitem.symbol;
            //    Costitem.cost = Costitem.cost;
            //    Pctotalcost += Costitem.packsAmnt * Costitem.cost;
            //}
            //PcTotaltb.Text = Pctotalcost.ToString();
            //finalCosttb.Text = Math.Round((newtotalcost + CostAdds + Pctotalcost),2).ToString();
            finalCosttb.Text = Math.Round((newtotalcost + CostAdds), 2).ToString();
            totalCosttb.Text = Math.Round(newtotalcost, 2).ToString();

            //double hrstotal = double.Parse(Esthrstotaltb.Text);
            //double hrscost = hrstotal * hrrate;

            //if (hrscost > 0)
            //{
            //AddToCosting(0,0, hrstotal, "", "Hours cost", hrrate, hrscost); 
            //HoursCosttb.Text = hrscost.ToString();
            //HoursCosttb.Text = "0";
            //Populate();
            // }
        }
        //private void DisplaySummary()
        //{
        //    foreach (string Str in SummaryList)
        //    {
        //        this.SummaryLB.Items.Add(Str);
        //    }
        //}
        private void CalculateHours()
        {
            HoursList.Clear();
            this.HoursLb.Items.Clear();
            double hoursTotal = 0;
            //
            if (coveList.Count != 0)
            {
                //double hrs = 0;
                foreach (Pack Pitem in coveList)
                {
                    if (Pitem.Options.CovingType == "Flake")
                    {
                        double FlakeCove = newBackend.lists.HrRates.FlakeCove;
                        hoursTotal += FlakeCove * Pitem.Meterage;
                        AddtoHoursSummary(Pitem.Name + " Flake Cove: " + FlakeCove * Pitem.Meterage + " Hrs");
                    }
                    if (Pitem.Options.CovingType == "50mm")
                    {
                        double Cove50 = newBackend.lists.HrRates.coving50mm;
                        hoursTotal += Cove50 * Pitem.Meterage;
                        AddtoHoursSummary(Pitem.Name + " 50mm Coves + seal: " + Cove50 * Pitem.Meterage + " Hrs");
                    }
                    if (Pitem.Options.CovingType == "75mm")
                    {
                        double Cove75 = newBackend.lists.HrRates.coving75mm;
                        hoursTotal += Cove75 * Pitem.Meterage;
                        AddtoHoursSummary(Pitem.Name + " 75mm Coves + seal: " + Cove75 * Pitem.Meterage + " Hrs");
                    }
                }
            }

            if (chevronList.Count != 0)
            {
                double hrs = 0;
                double chevr = newBackend.lists.HrRates.Chevron;
                foreach (Pack pItem in chevronList)
                {
                    hrs += chevr * pItem.Meterage;
                }
                hoursTotal += hrs;
                AddtoHoursSummary("Chevron: " + hrs + " Hrs");
            }

            foreach (Pack paItem in PUCfloorList)
            {
                double griding1 = newBackend.lists.HrRates.Grind1;
                double griding2 = newBackend.lists.HrRates.Grind2;
                if (paItem.Options.Grind == "Single Pass")
                {
                    hoursTotal += griding1 * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " Grinding Single Pass: " + griding1 * paItem.Meterage + " Hrs");
                }
                if (paItem.Options.Grind == "Double Pass")
                {
                    hoursTotal += griding2 * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " Grinding Double Pass: " + griding2 * paItem.Meterage + " Hrs");
                }

                double rc1 = newBackend.lists.HrRates.RollCprep1;
                double rc2 = newBackend.lists.HrRates.RollCprep2;
                double rc3 = newBackend.lists.HrRates.RollCprep3;
                if (paItem.Options.RollCoat == "RollCoat + Prep 1 Coat")
                {
                    hoursTotal += rc1 * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " RollCoat + Prep 1 Coat: " + rc1 * paItem.Meterage + " Hrs");
                }
                if (paItem.Options.RollCoat == "RollCoat + Prep 2 Coats")
                {
                    hoursTotal += rc2 * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " RollCoat + Prep 2 Coats: " + rc2 * paItem.Meterage + " Hrs");
                }
                if (paItem.Options.RollCoat == "RollCoat + Prep 3 Coats")
                {
                    hoursTotal += rc3 * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " RollCoat + Prep 3 Coats: " + rc3 * paItem.Meterage + " Hrs");
                }

                double tr4mm = newBackend.lists.HrRates.Trowel4mm;
                double tr4t6mmns = newBackend.lists.HrRates.Trowel6mm;
                if (paItem.Options.Trowel == "4mm Trowel Without BroadCast")
                {
                    hoursTotal += tr4mm * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " 4mm Trowel Without BroadCast: " + tr4mm * paItem.Meterage + " Hrs");
                }
                if (paItem.Options.Trowel == "4mm-6mm Trowel With Broadcast")
                {
                    hoursTotal += tr4t6mmns * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " 4mm-6mm Trowel With Broadcast: " + tr4t6mmns * paItem.Meterage + " Hrs");
                }

                if (paItem.Options.Polivac > 0)
                {

                    double PolyV = newBackend.lists.HrRates.Polyvac;
                    hoursTotal += PolyV * paItem.Meterage;
                    AddtoHoursSummary(paItem.Name + " Polyvac: " + PolyV * paItem.Meterage + " Hrs");
                }
            }

            if (linemarkList.Count != 0)
            {
                double totalMeters = 0;
                foreach (Pack pItm in linemarkList)
                {
                    totalMeters += pItm.Meterage;
                }
                double LineM = newBackend.lists.HrRates.Linemark;
                double hrs = 0;
                hrs += LineM * totalMeters;
                hoursTotal += hrs;
                AddtoHoursSummary("Line Marking: " + hrs + " Hrs");
                // }
            }

            if (subfillList.Count != 0)
            {
                double SubF = newBackend.lists.HrRates.Subfill;
                double hrs = 0;
                foreach (Pack Pitem in subfillList)
                {
                    hrs += SubF * Pitem.Meterage;
                }
                hoursTotal += hrs;
                AddtoHoursSummary("SubFill: " + hrs + " Hrs");
            }

            if (PCList.Count != 0)
            {
                double Gns = newBackend.lists.HrRates.PcGnc;
                double Mp = newBackend.lists.HrRates.pcMp;
                double EdgeGns = newBackend.lists.HrRates.PcEdgeGnc;
                double EdgeMp = newBackend.lists.HrRates.pcEdgeMp;
                double hrs = 0;
                foreach (Pack Pitem in PCList)
                {
                    if (Pitem.Name.Contains("Grind & Seal"))
                    {
                        if (Pitem.Name.Contains("Light"))
                        {
                            double Area = Pitem.Meterage;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = Pitem.Coverage / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;

                            AddtoHoursSummary("Grind & Seal-Light: " + hrs + " Hrs");
                        }
                        if (Pitem.Name.Contains("Medium"))
                        {
                            double Area = Pitem.Meterage * 1.262;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.262) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;
                            AddtoHoursSummary("Grind & Seal-Medium: " + hrs + " Hrs");
                        }
                        if (Pitem.Name.Contains("Full"))
                        {
                            double Area = Pitem.Meterage * 1.54;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.54) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;

                            AddtoHoursSummary("Grind & Seal-Full: " + hrs + " Hrs");
                        }
                    }
                    if (Pitem.Name.Contains("Mechanical Polish"))
                    {
                        if (Pitem.Name.Contains("Light"))
                        {

                            double Area = Pitem.Meterage;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;

                            AddtoHoursSummary("Mechanical Polish-Light: " + hrs + " Hrs");
                        }
                        if (Pitem.Name.Contains("Medium"))
                        {
                            double Area = Pitem.Meterage * 1.235;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.235) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;
                            AddtoHoursSummary("Mechanical Polish-Medium: " + hrs + " Hrs");
                        }
                        if (Pitem.Name.Contains("Full"))
                        {
                            double Area = Pitem.Meterage * 2;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 2) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;
                            AddtoHoursSummary("Mechanical Polish-Full: " + hrs + " Hrs");
                        }
                    }
                    //hoursTotal += hrs;
                    //AddtoHoursSummary("SubFill: " + hrs + " Hrs");
                }
            }

            double hrrate = 0;
            double hrstotal = 0;
            if (this.MainHrRatetb.Text != "")
                hrrate = double.Parse(this.MainHrRatetb.Text);
            if (this.Additiontb.Text != "")
                hrstotal = (hoursTotal + double.Parse(Additiontb.Text));
            double hrscost = hrstotal * hrrate;

            if (hrscost > 0)
            {
                //AddToCosting(0,0, hrstotal, "", "Hours cost", hrrate, hrscost); 
                HoursCosttb.Text = hrscost.ToString();
            }
            //this.TotalHrsTb.Text = hoursTotal.ToString();
            DisplayHours();
        }
        private void AddtoHoursSummary(string Instr)
        {
            HoursList.Add(Instr);
        }
        private void DisplayHours()
        {
            foreach (string Str in HoursList)
            {
                this.HoursLb.Items.Add(Str);
            }
        }
        private void AddToCartBtn_Click(object sender, EventArgs e)
        {
            AddtoCart();
        }

        private void AddtoCart()
        {
            string floortype = "";
            int Selected = 0;
            if (FloorSysLv.SelectedIndices.Count > 0)
            {
                Selected = FloorSysLv.SelectedIndices[0];
                string item = FloorSysLv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        floortype = item;
                        break;
                    }
                }
            }

            //string floortype = this.FloorType_Cb.Text;

            double Area = 0;
            bool isNumeric = double.TryParse(this.FsAreaTb.Text, out Area);
            if (floortype != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == floortype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Area;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            PUCfloorList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }
        public void UpdateConsumesBox()
        {
            this.Conslv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Consumable")
                {
                    string[] newArr = new string[12];
                    if (val.BaseMat.ToUpper().Contains(ConsumesSearchtb.Text.ToUpper()))
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        Conslv.Items.Add(itm);
                    }
                    else if (ConsumesSearchtb.Text == "")
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        Conslv.Items.Add(itm);
                    }
                }
            }
        }

        public void UpdateFloorsysBox()
        {
            this.FloorSysLv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Floor-PUC")
                {
                    string[] newArr = new string[12];
                    if (val.BaseMat.ToUpper().Contains(FsSearchtb.Text.ToUpper()))
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorSysLv.Items.Add(itm);
                    }
                    else if (FsSearchtb.Text == "")
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorSysLv.Items.Add(itm);
                    }
                    //this.FloorType_Cb.Items.AddRange(new object[] { val.BaseMat });
                }
            }
        }
        public void UpdateRCFloorBox()
        {
            this.FloorRCLv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Floor-RC")
                {
                    string[] newArr = new string[12];
                    if (val.BaseMat.ToUpper().Contains(FsRCsearchtb.Text.ToUpper()))
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorRCLv.Items.Add(itm);
                    }
                    else if (FsESLsearchtb.Text == "")
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorRCLv.Items.Add(itm);
                    }
                }
            }
        }

        public void UpdateBCFloorBox()
        {
            this.FloorBCFLv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Floor-BCF")
                {
                    string[] newArr = new string[12];
                    if (val.BaseMat.ToUpper().Contains(FsBCFsearchtb.Text.ToUpper()))
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorBCFLv.Items.Add(itm);
                    }
                    else if (FsBCFsearchtb.Text == "")
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorBCFLv.Items.Add(itm);
                    }
                }
            }
        }

        public void UpdateESLloorBox()
        {
            this.FloorESLLv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Floor-ESL")
                {
                    string[] newArr = new string[12];
                    if (val.BaseMat.ToUpper().Contains(FsESLsearchtb.Text.ToUpper()))
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorESLLv.Items.Add(itm);
                    }
                    else if (FsESLsearchtb.Text == "")
                    {
                        newArr[0] = val.BaseMat;
                        newArr[1] = val.PackList.Count.ToString();
                        ListViewItem itm = new ListViewItem(newArr);
                        FloorESLLv.Items.Add(itm);
                    }
                }
            }
        }
        public void UpdatePrepsysBox()
        {
            this.Prepsyslv.Items.Clear();
            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                if (val.Type == "Preperation")
                {
                    string[] newArr = new string[12];
                    newArr[0] = val.BaseMat;
                    newArr[1] = val.PackList.Count.ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    Prepsyslv.Items.Add(itm);
                }
            }
        }
        public void updateQuotebox()
        {
            Quoteslv.Items.Clear();
            foreach (Quote_Block Qfsitem in newBackend.lists.Q_bList)
            {
                string[] newArr = new string[40];
                if (Qfsitem.QName.ToUpper().Contains(QuotesTextbox.Text.ToUpper()))
                {
                    newArr[0] = Qfsitem.QName;
                    ListViewItem itm = new ListViewItem(newArr);
                    Quoteslv.Items.Add(itm);
                }
                else if (QuotesTextbox.Text == "")
                {
                    newArr[0] = Qfsitem.QName;
                    ListViewItem itm = new ListViewItem(newArr);
                    Quoteslv.Items.Add(itm);
                }
            }
        }

        public void updateBoxes()
        {
            //this.FloorType_Cb.Items.Clear();
            this.CovProdCb.Items.Clear();
            this.LineMarkcb.Items.Clear();
            this.Chevroncb.Items.Clear();
            this.Subfillcb.Items.Clear();
            this.PcFloorSyscb.Items.Clear();
            this.JointSyscb.Items.Clear();

            foreach (ProductSystem val in newBackend.lists.SystemsList)
            {
                //if (val.Type == "Floor")
                //{
                //    //FloorSysLv.Items.Add(val.BaseMat);
                //    //this.FloorType_Cb.Items.AddRange(new object[] { val.BaseMat });
                //}
                if (val.Type == "Coving")
                    this.CovProdCb.Items.AddRange(new object[] { val.BaseMat });
                if (val.Type == "LineMarking")
                    this.LineMarkcb.Items.AddRange(new object[] { val.BaseMat });
                if (val.Type == "Chevron")
                    this.Chevroncb.Items.AddRange(new object[] { val.BaseMat });
                if (val.Type == "Subfill")
                    this.Subfillcb.Items.AddRange(new object[] { val.BaseMat });
                if (val.Type == "Polished-Concrete")
                    this.PcFloorSyscb.Items.AddRange(new object[] { val.BaseMat });
                if (val.Type == "Joints")
                    this.JointSyscb.Items.AddRange(new object[] { val.BaseMat });
            }

            this.Chevroncb.Items.AddRange(new object[] { "Line" });
            this.Chevroncb.Items.AddRange(new object[] { "Number" });
            this.Chevroncb.Items.AddRange(new object[] { "WheelChair" });
            this.Chevroncb.Items.AddRange(new object[] { "Keep Clear" });

            //HrRatetb.Text = newBackend.lists.HrRates.HrMulti.ToString();

        }
        public void populateQuotelists()
        {
            updateQuotebox();
            newBackend.lists.RebuildQuoteList();
            Quoteslv.Items.Clear();
            foreach (Quote_Block Qfsitem in newBackend.lists.Q_bList)
            {
                string[] newArr = new string[15];
                newArr[0] = Qfsitem.QName;
                ListViewItem itm = new ListViewItem(newArr);
                Quoteslv.Items.Add(itm);
            }
        }

        public void populateStage()
        {
            this.Stagelv.Items.Clear();
            StageList.Sort((x, y) => x.Staging.CompareTo(y.Staging));

            foreach (CSstage stageitem in StageList)
            {
                string[] newArr = new string[10];
                newArr[0] = stageitem.Staging;
                newArr[1] = stageitem.Systems;
                newArr[2] = stageitem.Area.ToString();
                newArr[3] = stageitem.Lm.ToString();
                double cstamnt = Math.Round(stageitem.packsAmnt, 3);
                newArr[4] = stageitem.inputname;
                newArr[5] = cstamnt.ToString();

                ListViewItem itm = new ListViewItem(newArr);
                this.Stagelv.Items.Add(itm);
            }

            //SummaryLB.Items.Clear();
            SummaryList1.Clear();
            List<string> Localstageslist = new List<string>();
            foreach (CSstage stgitem in StageList)
            {
                string currentstring = stgitem.Staging;
                if (!Localstageslist.Contains(currentstring))
                {
                    Localstageslist.Add(stgitem.Staging);
                }
            }
            Localstageslist.Sort();
            foreach (string stritem in Localstageslist)
            {
                double stagecost = 0;
                double Area = 0;
                SummaryOutput newsum = new SummaryOutput();
                newsum.F1 = "----------------";
                newsum.F2 = stritem;
                newsum.F3 = "----------------";
                SummaryList1.Add(newsum);
                bool Warning = false;
                foreach (CSstage stgitem in StageList)
                {
                    if (stgitem.Staging == stritem)
                    {
                        newsum = new SummaryOutput();
                        newsum.F1 = stgitem.inputname;
                        if (stgitem.symbol == "Hour")
                            newsum.F2 = "$" + Math.Round((stgitem.price / stgitem.Area) * double.Parse("1." + CostMarkuptb.Text), 2) + " per " + stgitem.symbol;
                        else
                            newsum.F2 = "$" + Math.Round((stgitem.price / stgitem.Area) * double.Parse("1." + CostMarkuptb.Text), 2) + " per m2";

                        newsum.F3 = "$" + Math.Round(stgitem.price * double.Parse("1." + CostMarkuptb.Text), 2);
                        SummaryList1.Add(newsum);
                        if (Area != 0 && Area != stgitem.Area && !Warning)
                        {
                            MessageBox.Show(stritem + " has different Areas, Consider adding systems to serperate stages if areas are diffrent", "Message,", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Warning = true;
                        }
                        else
                        {
                            Area = stgitem.Area;

                        }
                        stagecost += Math.Round(stgitem.price, 2);
                    }
                }

                newsum = new SummaryOutput();
                newsum.F1 = "----------------";
                newsum.F2 = stritem + " Area = " + Math.Round(Area, 2);
                newsum.F3 = "----------------";
                SummaryList1.Add(newsum);

                newsum = new SummaryOutput();
                newsum.F1 = "----------------";
                newsum.F2 = stritem + " Price $" + Math.Round(stagecost * double.Parse("1." + CostMarkuptb.Text), 2);
                newsum.F3 = "----------------";
                SummaryList1.Add(newsum);

                newsum = new SummaryOutput();
                newsum.F1 = "----------------";
                newsum.F2 = stritem + " Price per M $" + Math.Round(stagecost * double.Parse("1." + CostMarkuptb.Text) / Area, 2);
                newsum.F3 = "----------------";
                SummaryList1.Add(newsum);

                newsum = new SummaryOutput();
                newsum.F1 = "";
                newsum.F2 = "";
                newsum.F3 = "";
                SummaryList1.Add(newsum);
            }
            cocalc();
            PopulateSummaryLv();
        }
        public void PopulateSummaryLv()
        {
            SummaryLv.Items.Clear();
            foreach (SummaryOutput Sitem in SummaryList1)
            {
                string[] newArr = new string[3];
                newArr[0] = Sitem.F1;
                newArr[1] = Sitem.F2;
                newArr[2] = Sitem.F3;
                ListViewItem itm = new ListViewItem(newArr);
                this.SummaryLv.Items.Add(itm);
            }
        }
        public void PopulatePrepList()
        {
            this.PrepListlv.Items.Clear();
            foreach (Pack Pitem in PrepList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.PrepListlv.Items.Add(itm);
            }
        }

        public void PopulateConsList()
        {
            this.Consumelv.Items.Clear();
            foreach (Pack Pitem in ConsumesList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.Consumelv.Items.Add(itm);
            }
        }
        public void populateAdds()
        {
            Addslv.Items.Clear();
            double Additionstotal = 0;
            foreach (Adds Additem in newBackend.lists.AddsList)
            {
                string[] newArr = new string[12];
                newArr[0] = Additem.ItemName;
                newArr[1] = Additem.Desc;
                newArr[2] = Additem.Ammount;
                Additionstotal += double.Parse(Additem.Cost);
                newArr[3] = Additem.Cost;
                ListViewItem itm = new ListViewItem(newArr);
                this.Addslv.Items.Add(itm);
            }
            CostAddstb.Text = Additionstotal.ToString();

        }
        public void populatAdditions()
        {
            this.Additionslv.Items.Clear();
            foreach (string stritem in newBackend.lists.AdditionsList)
            {
                this.Additionslv.Items.Add(stritem);
            }
        }
        public void Populate()
        {
            updateBoxes();
            this.FloorLv.Items.Clear();
            foreach (Pack Pitem in PUCfloorList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.FloorLv.Items.Add(itm);
            }

            this.RCLv.Items.Clear();
            foreach (Pack Pitem in RCfloorList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.RCLv.Items.Add(itm);
            }

            this.BCFLv.Items.Clear();
            foreach (Pack Pitem in BCFfloorList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.BCFLv.Items.Add(itm);
            }

            this.ESLLv.Items.Clear();
            foreach (Pack Pitem in ESLfloorList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.ESLLv.Items.Add(itm);
            }

            this.PolishedConclv.Items.Clear();
            foreach (Pack Pitem in PCList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.PriceItemList.Count.ToString();
                newArr[2] = Pitem.Meterage.ToString();
                newArr[3] = Pitem.Coverage.ToString() + " " + Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.PolishedConclv.Items.Add(itm);
            }

            this.Covelv.Items.Clear();
            foreach (Pack Pitem in coveList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Meterage.ToString();
                newArr[2] = Pitem.Price.ToString();
                newArr[3] = Pitem.Coverage.ToString();
                newArr[4] = Pitem.Unit;
                newArr[5] = Pitem.Options.CovingType;
                ListViewItem itm = new ListViewItem(newArr);
                this.Covelv.Items.Add(itm);
            }

            this.LineMarklv.Items.Clear();
            foreach (Pack Pitem in linemarkList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Meterage.ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.LineMarklv.Items.Add(itm);
            }

            this.JointsLv.Items.Clear();
            foreach (Pack Pitem in Jointlist)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Meterage.ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.JointsLv.Items.Add(itm);
            }

            this.ChevronLv.Items.Clear();
            foreach (Pack Pitem in chevronList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Meterage.ToString();
                newArr[2] = Pitem.Price.ToString();
                newArr[3] = Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.ChevronLv.Items.Add(itm);
            }

            this.SubfillLv.Items.Clear();
            foreach (Pack Pitem in subfillList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Meterage.ToString();
                newArr[2] = Pitem.Volume.ToString();
                newArr[3] = Pitem.Unit.ToString();
                newArr[4] = Pitem.Coverage.ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.SubfillLv.Items.Add(itm);
            }
        }
        private void ClearCartBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (FloorLv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < FloorLv.SelectedIndices.Count; i++)
                {
                    Selected = FloorLv.SelectedIndices[0];
                    PUCfloorList.RemoveAt(Selected);
                }
                Populate();
            }
        }
        private void AddCoveBtn_Click(object sender, EventArgs e)
        {
            string CoveType = this.CovProdCb.Text;
            double Length;
            bool isNumeric = double.TryParse(this.CoveLentb.Text, out Length);
            if (CoveType != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == CoveType)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Length;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            coveList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }
        private void ClearCoveBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Covelv.SelectedIndices.Count > 0)
            {
                Selected = Covelv.SelectedIndices[0];
                coveList.RemoveAt(Selected);
                Populate();
            }
        }
        private void LineMarkAddbtn_Click(object sender, EventArgs e)
        {
            string LineType = this.LineMarkcb.Text;
            double Length;
            bool isNumeric = double.TryParse(this.LineMarktb.Text, out Length);
            if (LineType != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == LineType)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            pItem.ParentSystemName = psItem.BaseMat;
                            pItem.Meterage = Length;
                            linemarkList.Add(pItem);
                        }
                    }
                }
            }
            Populate();
        }
        private void LineMarkClearbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (LineMarklv.SelectedIndices.Count > 0)
            {
                Selected = LineMarklv.SelectedIndices[0];
                linemarkList.RemoveAt(Selected);
                Populate();
            }
        }
        private void ResetOpbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            //int Selected2 = 0;
            if (ChevronLv.SelectedIndices.Count > 0)
            {
                Selected = ChevronLv.SelectedIndices[0];
                chevronList.RemoveAt(Selected);
                Populate();
            }
        }
        private void Chevroncb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Chevroncb.Text == "Line")
            {
                //Chevtb.Enabled = true;
                ChvPritb.Enabled = true;
                ChvPritb.Text = "10";
            }
            else
            {
                //Chevtb.Enabled = false;
                //Chevtb.Text = "";
                ChvPritb.Enabled = false;
                ChvPritb.Text = "";
            }
        }
        private void ChevronAddbrn_Click(object sender, EventArgs e)
        {
            string CHVtype = this.Chevroncb.Text;
            double Area;
            bool isNumeric = double.TryParse(this.Chevtb.Text, out Area);

            if (CHVtype != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == CHVtype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            pItem.ParentSystemName = psItem.BaseMat;
                            pItem.Meterage = Area;
                            chevronList.Add(pItem);
                        }
                    }
                }
            }
            Populate();


            if (CHVtype != "" && CHVtype == "Line" && isNumeric)
            {
                Priceitem prItem = new Priceitem();
                prItem.Name = ChvColourTb.Text + " Chevron Lines";
                prItem.Price = double.Parse(ChvPritb.Text);
                prItem.Weight = 0;
                prItem.Volume = 1;
                prItem.Coverage = 2;
                prItem.Meterage = Area;
                prItem.Unit = "m2/L";

                Pack pItem = new Pack();
                pItem.Name = "Chevron Lines";
                pItem.Price = double.Parse(ChvPritb.Text); ;
                pItem.PriceItemList.Add(prItem);
                pItem.Coverage = 2;
                pItem.Volume = 1;
                pItem.Weight = 0;
                pItem.Unit = "m";
                pItem.Meterage = Area;
                chevronList.Add(pItem);
            }
            else if (CHVtype != "")
            {
                if (CHVtype == "Number")
                {
                    Priceitem prItem = new Priceitem();
                    prItem.Name = ChvColourTb.Text + " Number";
                    prItem.Price = 14;
                    prItem.Weight = 0;
                    prItem.Volume = 1;
                    prItem.Coverage = 1;
                    prItem.Meterage = 1;
                    prItem.Unit = "Stencil";

                    Pack pItem = new Pack();
                    pItem.Name = "Number";
                    pItem.Price = 14;
                    pItem.PriceItemList.Add(prItem);
                    pItem.Coverage = 1;
                    pItem.Volume = 1;
                    pItem.Weight = 0;
                    pItem.Unit = "Item";
                    pItem.Meterage = 1;
                    chevronList.Add(pItem);
                }
                if (CHVtype == "WheelChair")
                {
                    Priceitem prItem = new Priceitem();
                    prItem.Name = ChvColourTb.Text + " WheelChair";
                    prItem.Price = 50;
                    prItem.Weight = 0;
                    prItem.Volume = 1;
                    prItem.Coverage = 1;
                    prItem.Meterage = 1;
                    prItem.Unit = "Stencil";

                    Pack pItem = new Pack();
                    pItem.Name = "WheelChair";
                    pItem.Price = 50;
                    pItem.PriceItemList.Add(prItem);
                    pItem.Coverage = 1;
                    pItem.Volume = 1;
                    pItem.Weight = 0;
                    pItem.Unit = "Item";
                    pItem.Meterage = 1;
                    chevronList.Add(pItem);
                }
                if (CHVtype == "Keep Clear")
                {
                    //PushToList("Keep Clear Chevron", 0.5, "L", 0, 1, 20, meterage, ChevronList, true);
                    Priceitem prItem = new Priceitem();
                    prItem.Name = ChvColourTb.Text + " Keep Clear";
                    prItem.Price = 20;
                    prItem.Weight = 0;
                    prItem.Volume = 1;
                    prItem.Coverage = 1;
                    prItem.Meterage = 1;
                    prItem.Unit = "Stencil";

                    Pack pItem = new Pack();
                    pItem.Name = "Keep Clear";
                    pItem.Price = 20;
                    pItem.PriceItemList.Add(prItem);
                    pItem.Coverage = 1;
                    pItem.Volume = 1;
                    pItem.Weight = 0;
                    pItem.Unit = "Item";
                    pItem.Meterage = 1;
                    chevronList.Add(pItem);
                }
            }
            Populate();
        }
        private void SubFilladdbtn_Click(object sender, EventArgs e)
        {
            string SFType = this.Subfillcb.Text;
            double Area;
            bool isNumeric = double.TryParse(this.SubFilltb.Text, out Area);
            double thick;
            bool ATisNumeric = double.TryParse(this.Avgthicktb.Text, out thick);
            if (SFType != "" && isNumeric && ATisNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == SFType)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = (Area * thick) / 1000;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            subfillList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }
        private void optionAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //SystemEdit f2 = new SystemEdit(newBackend); //this is the change, code for redirect
            //f2.Owner = this;
            //f2.Show();
            // this.Hide();
        }

        private void productsAndProductsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ProductEdit f3 = new ProductEdit(newBackend);
            //f3.Owner = this;
            //f3.Show();
            //this.Hide();
        }

        public void Reload()
        {
            newBackend.lists.ClearLists();
            //newBackend.lists.LoadFSData();
            //newBackend.lists.LoadPackData();
            //newBackend.lists.LoadProductData();
            Populate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (SubfillLv.SelectedIndices.Count > 0)
            {
                Selected = SubfillLv.SelectedIndices[0];
                subfillList.RemoveAt(Selected);
                Populate();
            }
        }

        private void hourRatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hrAddbtn form2 = new hrAddbtn(newBackend);
            form2.Owner = this;
            form2.Show();
            //this.Hide();
        }

        private void FSclearbtn_Click(object sender, EventArgs e)
        {
            PUCfloorList.Clear();
            //FloorType_Cb.Text = "";
            FsAreaTb.Text = "";
            Populate();
        }

        private void Covclearbtn_Click(object sender, EventArgs e)
        {
            coveList.Clear();
            CovProdCb.Text = "";
            CoveLentb.Text = "";
            Populate();
        }

        private void SFclearbtn_Click(object sender, EventArgs e)
        {
            subfillList.Clear();
            Subfillcb.Text = "";
            SubFilltb.Text = "";
            Avgthicktb.Text = "";
            Populate();
        }

        private void LMclearbtn_Click(object sender, EventArgs e)
        {
            linemarkList.Clear();
            LineMarkcb.Text = "";
            LineMarktb.Text = "";
            Populate();
        }

        private void ChvClearbtn_Click(object sender, EventArgs e)
        {
            chevronList.Clear();
            ChvColourTb.Text = "";
            Chevroncb.Text = "";
            Chevtb.Text = "";
            Populate();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //int Selected = 0;
            //if (FloorLv.SelectedIndices.Count > 0)
            //{
            //    Selected = FloorLv.SelectedIndices[0];
            //    string itemName = FloorLv.Items[Selected].Text;
            //    foreach (Pack PsItem in floorList)
            //    {
            //        if (PsItem.Name == itemName)
            //        {
            //            PsItem.Options.Grind = Grcb.Text;
            //            PsItem.Options.RollCoat = RCcb.Text;
            //            PsItem.Options.Trowel = TRcb.Text;
            //            if (Polyvaccb.Checked)
            //                PsItem.Options.Polivac = int.Parse(PoliPassestb.Text);
            //        }
            //    }
            //}
            //Populate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Covelv.SelectedIndices.Count > 0)
            {
                Selected = Covelv.SelectedIndices[0];
                string itemName = Covelv.Items[Selected].Text;
            }
            Populate();
        }


        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Info inform = new Info();
            inform.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.qepoxy.com.au/");
        }

        private void allInOneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AIOne f3 = new AIOne(newBackend);
            f3.Owner = this;
            f3.Show();
        }


        private void jobGenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Jobnum f4 = new Jobnum();
            f4.Owner = this;
            f4.Show();
        }

        private void quotationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double FinCost = 0;
            bool isNumeric = double.TryParse(finalCosttb.Text, out FinCost);
            //finalCosttb.Text
            if (isNumeric)
            {
                Quote QuoteForm = new Quote(newBackend, FinCost, QuotesTextbox.Text);
                QuoteForm.Show();
            }
            else
            {
                Quote QuoteForm = new Quote(newBackend, double.Parse(totalCosttb.Text), QuotesTextbox.Text);
                QuoteForm.Show();
            }
        }

        private void purchaseOrderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PurchaseOrder PoExporter = new PurchaseOrder(newBackend);
            //PoExporter.Owner = this;
            PoExporter.Show();
        }

        private void budgetToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private bool errorCheckName(string inputErr)
        {
            bool error = false;
            if (inputErr.Contains(","))
            {
                error = true;
            }
            return error;
        }
        private void QsaveBtn_Click(object sender, EventArgs e)
        {
            bool noError = true;

            if (QuotesTextbox.Text != "")
            {
                if (errorCheckName(QuotesTextbox.Text))
                {
                    this.errorProvider1.SetError(this.QuotesTextbox, "contains ',' cannot save due to bad charecter");
                    noError = false;
                }
                if (noError)
                {
                    this.errorProvider1.Clear();
                    Quote_Block Qb_item = new Quote_Block();
                    Qb_item.QName = QuotesTextbox.Text;
                    Qb_item.savedHradd = Additiontb.Text;
                    Qb_item.savedHrRate = MainHrRatetb.Text;
                    Qb_item.savedAdds = CostAddstb.Text;
                    Qb_item.savedMarkup = CostMarkuptb.Text;
                    string today = DateTime.Now.ToString("dd/MM/yyyy");
                    Qb_item.SaveDate = today;

                    foreach (CSstage csitem in StageList)
                    {
                        Qb_item.qbStagelist.Add(csitem);
                    }

                    foreach (Adds additem in newBackend.lists.AddsList)
                    {
                        Qb_item.qbAdditionslist.Add(additem);
                    }

                    double count = EstHrsDgv.Rows.Count;

                    for (int i = 0; i < count - 1; i++)
                    {
                        EstHoursSave estsave = new EstHoursSave();
                        estsave.m2 = EstHrsDgv.Rows[i].Cells[0].Value.ToString();
                        estsave.lm = EstHrsDgv.Rows[i].Cells[1].Value.ToString();
                        estsave.stageName = EstHrsDgv.Rows[i].Cells[2].Value.ToString();
                        estsave.systemName = EstHrsDgv.Rows[i].Cells[3].Value.ToString();
                        estsave.method = EstHrsDgv.Rows[i].Cells[4].Value.ToString();
                        estsave.techHours = "";

                        if (EstHrsDgv.Rows[i].Cells[5].Value != null)
                        {
                            estsave.techHours = EstHrsDgv.Rows[i].Cells[5].Value.ToString();//tech hrs
                        }
                        estsave.Hours = EstHrsDgv.Rows[i].Cells[6].Value.ToString();//calc hrs
                        Qb_item.estHrslist.Add(estsave);
                    }
                    newBackend.lists.AddtoQuote_BlockList(Qb_item);
                    updateBox("Data Saved");
                }
            }
            populateQuotelists();
        }

        private void Q_bloadBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Quoteslv.SelectedIndices.Count > 0)
            {
                Selected = Quoteslv.SelectedIndices[0];
                string item = Quoteslv.Items[Selected].Text;
                QuotesTextbox.Text = item;
            }

            if (QuotesTextbox.Text != "")
            {
                string Updatestr = "";
                foreach (Quote_Block Qbitem in newBackend.lists.Q_bList)
                {
                    if (Qbitem.QName == QuotesTextbox.Text)
                    {
                        //QFs.Name = QuoteNumcb.Text;
                        MainHrRatetb.Text = Qbitem.savedHrRate;
                        Additiontb.Text = Qbitem.savedHradd;
                        CostAddstb.Text = Qbitem.savedAdds;
                        CostMarkuptb.Text = Qbitem.savedMarkup;
                        Updatestr = "Quote Save Date: " + Qbitem.SaveDate + "\n" + "Hour Rate set to: " + MainHrRatetb.Text + "\n" + "Cost Additions set to: " + Additiontb.Text + "\n" + "Cost Markup set to: " + CostMarkuptb.Text + "\n" + "Additions set to: " + Additiontb.Text;
                    }

                    if (Qbitem.QName == QuotesTextbox.Text)
                    {
                        StageList.Clear();
                        foreach (CSstage csitem in Qbitem.qbStagelist)
                        {
                            StageList.Add(csitem);
                        }
                        EstHrsDgv.Rows.Clear();

                        foreach (EstHoursSave hritem in Qbitem.estHrslist)
                        {
                            EstHrsDgv.Rows.Add(hritem.m2, hritem.lm, hritem.stageName, hritem.systemName, hritem.method, hritem.techHours, hritem.Hours);
                        }
                        newBackend.lists.AddsList.Clear();
                        foreach (Adds addsitem in Qbitem.qbAdditionslist)
                        {
                            newBackend.lists.AddsList.Add(addsitem);
                        }

                    }
                }
                populateAdds();
                populateStage();
                populateHrrates();
                PopulateHrs();
                cocalc();
                importToHrsList();
                //CalcBtn_Click(sender, e);
                updateBox(Updatestr);
            }
        }
        private void QremBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (Quoteslv.SelectedIndices.Count > 0)
            {
                Selected = Quoteslv.SelectedIndices[0];
                item = Quoteslv.Items[Selected].Text;
            }

            if (item != "")
            {
                updateBox(QuotesTextbox.Text + " Has Been Deleted");
                newBackend.lists.DelQuote(item);
                //newBackend.lists.RebuildQuoteList();
                populateQuotelists();
            }
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {
            if (QuotesTextbox.Text != "" && CostingList.Count > 0)
            {
                string SDate = "";
                foreach (QuoteFS Qfs in newBackend.lists.QuoteFSdata)
                {
                    if (Qfs.Name == QuotesTextbox.Text)
                    {
                        SDate = Qfs.SaveDate;
                    }
                }

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                object start_of_doc = "\\startofdoc";
                string[,] CstListArray;
                List<Costing> cstList = new List<Costing>();
                cstList.Clear();
                CstListArray = new string[costList.Count + 2, 5];
                int i = 1;
                foreach (Costing Coitem in CostingList)
                {
                    i++;
                    cstList.Add(Coitem);
                    CstListArray[i, 1] = Coitem.inputname;
                    CstListArray[i, 2] = Math.Round(Coitem.Ammount, 2).ToString() + " " + Coitem.symbol;
                    CstListArray[i, 3] = "$" + Math.Round(Coitem.cost, 2).ToString();
                    CstListArray[i, 4] = "$" + Math.Round(Coitem.price, 2).ToString();
                }

                //Start Word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;
                oWord = new Word.Application();
                oWord.Visible = true;
                oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

                //Word.Paragraph oHead1;

                Object oMissed = oDoc.Paragraphs[1].Range;
                Object oLinkToFile = false;
                Object oSaveWithDocument = true;
                //oDoc.InlineShapes.AddPicture(fileName, ref oLinkToFile, ref oSaveWithDocument, ref oMissed);


                Word.Paragraph oPara0;
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara0 = oDoc.Content.Paragraphs.Add(ref oRng);
                //oPara0.Range.Text = "Heading 2";
                oPara0.Format.SpaceAfter = 6;
                oPara0.Range.InsertParagraphAfter();
                //Insert Picture    
                Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

                string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
                string path = wanted_path + "\\Data\\Qepoxy1.jpg";
                range.Application.ActiveDocument.InlineShapes.AddPicture(@path, ref oLinkToFile, ref oSaveWithDocument, ref oMissed);

                //Microsoft.Office.Interop.Word.Shape s = oDoc.Shapes.AddShape((int)Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 320, 100, 250, 35);
                //s.Fill.ForeColor.RGB = Color.FromArgb(81, 208, 146).ToArgb();
                //s.TextFrame.TextRange.Text = "Purchase Order";
                //s.TextFrame.TextRange.Font.Size = 24;
                //s.TextFrame.TextRange.Font.Name = "Arial";
                //s.TextFrame.TextRange.Font.Color = 0;

                foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
                {
                    Word.Range header = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    Word.Paragraph p1 = header.Paragraphs.Add(header);
                    header.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    header.Text = "Version: 2.1";
                    header.Font.Size = 8;
                    header.Font.Bold = 0;
                    // header.InsertParagraphAfter();
                }

                int spacing = 12;
                //Insert a paragraph at the beginning of the document.
                Word.Paragraph oPara1a;
                oPara1a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara1a.Range.Text = "Quote Number: " + QuotesTextbox.Text;
                oPara1a.Range.Font.Bold = 1;
                oPara1a.Range.Font.Size = 14;
                oPara1a.Range.Font.Name = "calibri";
                oPara1a.Format.SpaceAfter = spacing;
                oPara1a.Range.InsertParagraphAfter();


                //string today = DateTime.Now.ToString("dd/MM/yyyy");
                //this.textBox2.Text = today;

                //Insert a paragraph at the beginning of the document.

                Word.Paragraph oPara1b;
                oPara1b = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara1b.Range.Text = "Date: " + SDate;
                oPara1b.Range.Font.Bold = 1;
                oPara1b.Range.Font.Size = 14;
                oPara1b.Range.Font.Name = "calibri";
                oPara1b.Format.SpaceAfter = spacing;
                oPara1b.Range.InsertParagraphAfter();

                Word.Paragraph oPara2;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
                int tablerows = cstList.Count + 1;
                Word.Table firstTable = oDoc.Tables.Add(oPara2.Range, tablerows, 4, ref oMissing, ref oMissing);

                firstTable.Borders.Enable = 0;
                foreach (Word.Row row in firstTable.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        //Header row  
                        if (cell.RowIndex == 1 && cell.ColumnIndex == 1)
                        {
                            cell.Range.Text = "Description: ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = 12;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 2)
                        {
                            cell.Range.Text = "Qty: ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = 12;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 3)
                        {
                            cell.Range.Text = "Unit Price: ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = 12;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 4)
                        {
                            cell.Range.Text = "Amount:";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = 12;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex <= tablerows)
                        {
                            //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                            cell.Range.Font.Bold = 0;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = 10;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            cell.Range.Text = CstListArray[cell.RowIndex, cell.ColumnIndex];
                        }
                        else
                        {
                            //cell.Range.Text = "$" + (Subtotal * 1.1).ToString();
                            //cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            //cell.Range.Font.Name = "calibri";
                            //cell.Range.Font.Size = 12;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            // cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                }

                Word.Paragraph oPara3z;
                oPara3z = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3z.Format.SpaceAfter = spacing;
                oPara3z.Range.InsertParagraphAfter();

                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3a.Range.Text = "Materials Total: " + MatTotaltb.Text;
                oPara3a.Range.Font.Bold = 1;
                oPara3a.Range.Font.Size = 14;
                oPara3a.Range.Font.Name = "calibri";
                oPara3a.Format.SpaceAfter = spacing;
                oPara3a.Range.InsertParagraphAfter();

                Word.Paragraph oPara3b;
                oPara3b = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3b.Range.Text = "Additions Costs: " + CostAddstb.Text;
                oPara3b.Range.Font.Bold = 1;
                oPara3b.Range.Font.Size = 14;
                oPara3b.Range.Font.Name = "calibri";
                oPara3b.Format.SpaceAfter = spacing;
                oPara3b.Range.InsertParagraphAfter();

                Word.Paragraph oPara3c;
                oPara3c = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3c.Range.Text = "Hours: " + HoursCosttb.Text;
                oPara3c.Range.Font.Bold = 1;
                oPara3c.Range.Font.Size = 14;
                oPara3c.Range.Font.Name = "calibri";
                oPara3c.Format.SpaceAfter = spacing;
                oPara3c.Range.InsertParagraphAfter();

                Word.Paragraph oPara3d;
                oPara3d = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3d.Range.Text = "Total: " + finalCosttb.Text;
                oPara3d.Range.Font.Bold = 1;
                oPara3d.Range.Font.Size = 14;
                oPara3d.Range.Font.Name = "calibri";
                oPara3d.Format.SpaceAfter = spacing;
                oPara3d.Range.InsertParagraphAfter();
            }
        }

        private void ClearListsbtn_Click(object sender, EventArgs e)
        {
            PUCfloorList.Clear();
            coveList.Clear();
            subfillList.Clear();
            linemarkList.Clear();
            chevronList.Clear();
            PCList.Clear();

            Updateboxes();
            Additiontb.Text = "0";
            CostAddstb.Text = "0";
            //HoursCosttb.Text = "0";
            PcTotaltb.Text = "0";
            //CostMarkuptb.Text = newBackend.lists.HrRates.CostMarkUp.ToString();
            //CalcBtn_Click(sender, e);

            stagehrslist.Clear();
            StageList.Clear();
            PcCostList.Clear();
            chevronList.Clear();
            linemarkList.Clear();
            subfillList.Clear();
            coveList.Clear();
            PUCfloorList.Clear();
            RCfloorList.Clear();
            BCFfloorList.Clear();
            ESLfloorList.Clear();
            PCList.Clear();
            PrepList.Clear();
            Jointlist.Clear();
            stagehrslist.Clear();

            populateHrrates();
            populateStage();
            Populate();
            updateHRScombo();
            populateStage();


        }
        async public void updateBox(string Updatestr)
        {
            StatusLabel1.Text = Updatestr;
            await Task.Delay(5000);
            StatusLabel1.Text = "";
        }

        private void PcAddbtn_Click(object sender, EventArgs e)
        {
            string floortype = this.PcFloorSyscb.Text;
            double Area = 0;
            bool AreaisNumeric = double.TryParse(this.PcAreatb.Text, out Area);
            if (floortype != "" && AreaisNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == floortype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Area;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            pItem.ParentSystemName = psItem.BaseMat;
                            foreach (Priceitem pritem in pItem.PriceItemList)
                            {
                                newp.PriceItemList.Add(pritem);
                            }
                            PCList.Add(newp);
                        }
                    }
                }
            }
            if (PCvarcb.Text != "" && Exposurecb.Text != "" && AreaisNumeric)
            {
                Pack newp = new Pack();
                newp.Meterage = double.Parse(PcAreatb.Text);
                newp.Name = PCvarcb.Text + "-" + Exposurecb.Text;

                double GrnSealLight = 0;
                double GrnSealMed = 0;
                double GrnSealFull = 0;
                double MpSealLight = 0;
                double MpSealMed = 0;
                double MpSealFull = 0;

                foreach (Hrrates hritem in newBackend.lists.HrrateList)
                {
                    if (hritem.hrsname == "GrnSealLightpr")
                        GrnSealLight = hritem.hrsWeight;
                    if (hritem.hrsname == "GrnSealMedpr")
                        GrnSealMed = hritem.hrsWeight;
                    if (hritem.hrsname == "GrnSealFullpr")
                        GrnSealFull = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealLightpr")
                        MpSealLight = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealMedpr")
                        MpSealMed = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealFullpr")
                        MpSealFull = hritem.hrsWeight;
                }
                if (newp.Name.Contains("Grind & Seal"))
                {
                    if (newp.Name.Contains("Light"))
                        newp.Price = GrnSealLight;
                    if (newp.Name.Contains("Medium"))
                        newp.Price = GrnSealMed;
                    if (newp.Name.Contains("Full"))
                        newp.Price = GrnSealFull;
                }
                if (newp.Name.Contains("Mechanical Polish"))
                {
                    if (newp.Name.Contains("Light"))
                        newp.Price = MpSealLight;
                    if (newp.Name.Contains("Medium"))
                        newp.Price = MpSealMed;
                    if (newp.Name.Contains("Full"))
                        newp.Price = MpSealFull;
                }

                newp.Weight = 0;
                newp.Volume = 1;
                newp.Unit = "Sqm";
                if (PcEdgeLtb.Text == "")
                {
                    newp.Coverage = Math.Sqrt(newp.Meterage) * 4;
                }
                else
                {
                    newp.Coverage = double.Parse(PcEdgeLtb.Text);
                }

                PCList.Add(newp);
            }
            Populate();
        }

        private void PcRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (PolishedConclv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < PolishedConclv.SelectedIndices.Count; i++)
                {
                    Selected = PolishedConclv.SelectedIndices[0];
                    PCList.RemoveAt(Selected);
                }
                Populate();
            }
        }

        private void PcClearbtn_Click(object sender, EventArgs e)
        {
            PCList.Clear();
            PcFloorSyscb.Text = "";
            Exposurecb.Text = "";
            PCvarcb.Text = "";
            PcAreatb.Text = "";
            PcEdgeLtb.Text = "";
            Populate();
        }

        private void pickListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CostEdit Ce1 = new CostEdit(StageList, CostingList, newBackend);
            Ce1.Owner = this;
            Ce1.Show();
        }

        private void Searchbtn_Click(object sender, EventArgs e)
        {

        }

        private void ClearXbtn_Click(object sender, EventArgs e)
        {
            QuotesTextbox.Text = "";
            Populate();
        }
        public void BoxUpdate()
        {
            MainHrRatetb.Text = newBackend.lists.HrRates.HrMulti.ToString();
            CostMarkuptb.Text = newBackend.lists.HrRates.CostMarkUp.ToString();
        }
        private void Quoteslv_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = Quoteslv.HitTest(e.X, e.Y);
            ListViewItem item = info.Item;

            if (item != null)
            {
                this.QuotesTextbox.Text = item.Text;
            }
            //else
            //{
            //    this.listView1.SelectedItems.Clear();
            //    this.QuotesTextbox.Text = "No Item is Selected";
            //}
        }

        private void FsSearchbtn_Click(object sender, EventArgs e)
        {
            UpdateFloorsysBox();
        }

        private void FsSearchtb_TextChanged(object sender, EventArgs e)
        {
            UpdateFloorsysBox();
        }

        private void dataControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataControl datacontrolerform = new DataControl(newBackend);
            datacontrolerform.Owner = this;
            datacontrolerform.Show();
        }

        private void QuotesTextbox_TextChanged(object sender, EventArgs e)
        {
            updateQuotebox();
        }

        private void FSaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = PUCStagecb.Text;
            string floorstr = "Floorsystem-PUCement";
            Stagelv.Items.Clear();
            foreach (Pack pItem in PUCfloorList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(FsAreaTb.Text), 0, prItem.Weight, prItem.Meterage / pk,
                                        pItem.Meterage / pk * prItem.Weight,
                                        prItem.Unit, prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(FsAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk,
                                        (pItem.Meterage / pk) * prItem.Volume,
                                        prItem.Unit,
                                        prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(FsAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            // HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void StageClearbtn_Click(object sender, EventArgs e)
        {
            StageList.Clear();

            EstHrsDgv.Columns.Clear();

            this.EstHrsDgv.Columns.Add("Area", "m2");
            this.EstHrsDgv.Columns.Add("Length", "Lm");
            this.EstHrsDgv.Columns.Add("Stage", "Stage Name");
            this.EstHrsDgv.Columns.Add("SysName", "System Name");
            this.EstHrsDgv.Columns.Add("Method", "Method");
            this.EstHrsDgv.Columns.Add("Techs", "Tech Hrs");
            this.EstHrsDgv.Columns.Add("Techs", "Hours");
            this.EstHrsDgv.Columns[0].Width = 80;
            this.EstHrsDgv.Columns[1].Width = 80;
            this.EstHrsDgv.Columns[2].Width = 275;

            HrCombolv.Items.Clear();
            newBackend.lists.AddsList.Clear();
            CostingList.Clear();
            SummaryList1.Clear();
            stagehrslist.Clear();
            //SummaryLB.Items.Clear();
            HoursLb.Items.Clear();
            newBackend.lists.HrComboList.Clear();

            this.Additiontb.Text = "0";
            totalCosttb.Text = "0";
            CostAddstb.Text = "0";
            finalCosttb.Text = "0";
            HoursCosttb.Text = "0";
            MatTotaltb.Text = "0";
            TotalHrsTb.Text = "0";
            PcTotaltb.Text = "0";
            populateAdds();
            populateStage();
            Populate();
            PopulateHrs();
            populateStage();
            updateHRScombo();
        }

        private void StageCoveAddbtn_Click(object sender, EventArgs e)
        {
            string covstagestr = CovStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in coveList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(covstagestr, "Coving", 0, double.Parse(CoveLentb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(covstagestr, "Coving", 0, double.Parse(CoveLentb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(covstagestr, "Coving", 0, double.Parse(CoveLentb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void SFaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = sfStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in subfillList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, "Subfill", double.Parse(SubFilltb.Text), double.Parse(Avgthicktb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Subfill", double.Parse(SubFilltb.Text), double.Parse(Avgthicktb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Subfill", double.Parse(SubFilltb.Text), double.Parse(Avgthicktb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);

                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void LMaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = lmStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in linemarkList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, "LineMark", 0, double.Parse(LineMarktb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "LineMark", 0, double.Parse(LineMarktb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "LineMark", 0, double.Parse(LineMarktb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void CHVaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = chvStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in chevronList)
            {


                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, "Chevron", double.Parse(Chevtb.Text), 0, prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Chevron", double.Parse(Chevtb.Text), 0, prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Chevron", double.Parse(Chevtb.Text), 0, prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];

        }

        private void PCaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = pcStagecb.Text;
            Stagelv.Items.Clear();
            double PCPriceTotal = 0;

            if (PCList.Count != 0)
            {

                double GrnSealLight = 0;
                double GrnSealMed = 0;
                double GrnSealFull = 0;
                double MpSealLight = 0;
                double MpSealMed = 0;
                double MpSealFull = 0;
                double Gns = 0;
                double Mp = 0;
                double EdgeGns = 0;
                double EdgeMp = 0;

                foreach (Hrrates hritem in newBackend.lists.HrrateList)
                {
                    if (hritem.hrsname == "GrnSealLightpr")
                        GrnSealLight = hritem.hrsWeight;
                    if (hritem.hrsname == "GrnSealMedpr")
                        GrnSealMed = hritem.hrsWeight;
                    if (hritem.hrsname == "GrnSealFullpr")
                        GrnSealFull = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealLightpr")
                        MpSealLight = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealMedpr")
                        MpSealMed = hritem.hrsWeight;
                    if (hritem.hrsname == "MpSealFullpr")
                        MpSealFull = hritem.hrsWeight;
                    if (hritem.hrsname == "PcGns")
                        Gns = hritem.hrsWeight;
                    if (hritem.hrsname == "PcMp")
                        Mp = hritem.hrsWeight;
                    if (hritem.hrsname == "PcEdgeGns")
                        EdgeGns = hritem.hrsWeight;
                    if (hritem.hrsname == "PcEdgeMp")
                        EdgeMp = hritem.hrsWeight;
                }

                double hrs = 0;
                foreach (Pack Pitem in PCList)
                {
                    if (Pitem.Name.Contains("Grind & Seal"))
                    {
                        if (Pitem.Name.Contains("Light"))
                        {
                            double Area = Pitem.Meterage;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = Pitem.Coverage / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;
                            //AddtoHoursSummary("Grind & Seal-Light: " + hrs + " Hrs");

                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Grind & Seal-Light";
                            newHrCitem.Rate = Gns;
                            //newHrCitem.Meterage = 
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);

                        }
                        if (Pitem.Name.Contains("Medium"))
                        {
                            double Area = Pitem.Meterage * 1.262;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.262) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;
                            //AddtoHoursSummary("Grind & Seal-Medium");
                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Grind & Seal-Medium";
                            newHrCitem.Rate = Gns;
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);
                        }
                        if (Pitem.Name.Contains("Full"))
                        {
                            double Area = Pitem.Meterage * 1.54;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.54) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Gns;
                            double Rates2 = hardArea * EdgeGns;
                            hrs = Rates1 + Rates2;

                            //AddtoHoursSummary("Grind & Seal-Full: " + hrs + " Hrs");
                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Grind & Seal-Full";
                            newHrCitem.Rate = Gns;
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);
                        }
                    }

                    if (Pitem.Name.Contains("Mechanical Polish"))
                    {
                        if (Pitem.Name.Contains("Light"))
                        {

                            double Area = Pitem.Meterage;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;

                            //AddtoHoursSummary("Mechanical Polish-Light");
                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Mechanical Polish-Light";
                            newHrCitem.Rate = Gns;
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);
                        }
                        if (Pitem.Name.Contains("Medium"))
                        {
                            double Area = Pitem.Meterage * 1.235;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 1.235) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;
                            //AddtoHoursSummary("Mechanical Polish-Medium");
                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Mechanical Polish-Medium";
                            newHrCitem.Rate = Gns;
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);
                        }
                        if (Pitem.Name.Contains("Full"))
                        {
                            double Area = Pitem.Meterage * 2;
                            double Perim = Math.Sqrt(Area) * 4;
                            double compRatio = (Pitem.Coverage * 2) / Perim;
                            double easyArea = Area - ((Perim * compRatio) / 2);
                            double hardArea = Area - easyArea;
                            double Rates1 = easyArea * Mp;
                            double Rates2 = hardArea * EdgeMp;
                            hrs = Rates1 + Rates2;
                            //AddtoHoursSummary("Mechanical Polish-Full: " + hrs + " Hrs");
                            HrCombo newHrCitem = new HrCombo();
                            newHrCitem.Stage = stagestr + " Polished Concrete";
                            newHrCitem.Method = "Mechanical Polish-Full";
                            newHrCitem.Rate = Gns;
                            double Lmin = double.Parse(PcEdgeLtb.Text);
                            double Areain = Area;

                            newHrCitem.Result = hrs;
                            newBackend.lists.HrComboList.Add(newHrCitem);
                        }
                    }
                }
            }
            updateHRScombo();
            DisplayHours();
            foreach (Pack pItem in PCList)
            {
                if (pItem.Name.Contains("Grind & Seal") || pItem.Name.Contains("Mechanical Polish"))
                {
                    double pk = 0;
                    Priceitem prItem = new Priceitem();
                    pk = pItem.Volume * pItem.Coverage;
                    AddToStage(stagestr, "Polish Concrete", double.Parse(PcAreatb.Text), double.Parse(PcEdgeLtb.Text), prItem.Volume, pItem.Meterage, pItem.Meterage * pItem.Volume, pItem.Unit, pItem.Name, pItem.Price, pItem.Meterage * pItem.Price, pItem.ParentSystemName);
                    PCPriceTotal += pItem.Meterage * pItem.Price;
                    //costList.Add(prItem);
                }
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, "Polish Concrete", double.Parse(PcAreatb.Text), double.Parse(PcEdgeLtb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Polish Concrete", double.Parse(PcAreatb.Text), double.Parse(PcEdgeLtb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Polish Concrete", double.Parse(PcAreatb.Text), double.Parse(PcEdgeLtb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void PrepaddtoStage()
        {
            string stagestr = PrepStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in PrepList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        if (Pareacb.Checked)
                            AddToStage(stagestr, "Preperation", double.Parse(PrepAreatb.Text), 0, prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        else if (Parealmcb.Checked)
                            AddToStage(stagestr, "Preperation", 0, double.Parse(PrepAreatb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        if (Pareacb.Checked)
                            AddToStage(stagestr, "Preperation", double.Parse(PrepAreatb.Text), 0, prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        else if (Parealmcb.Checked)
                            AddToStage(stagestr, "Preperation", 0, double.Parse(PrepAreatb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        if (Pareacb.Checked)
                            AddToStage(stagestr, "Preperation", double.Parse(PrepAreatb.Text), 0, prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        else if (Parealmcb.Checked)
                            AddToStage(stagestr, "Preperation", 0, double.Parse(PrepAreatb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
        }
        private void ConsaddtoStage()
        {
            string stagestr = ConsAtoScb.Text;
            Stagelv.Items.Clear();
            if (Areamulticb.Checked)
            {
                foreach (Pack pItem in ConsumesList)
                {
                    foreach (Priceitem prItem in pItem.PriceItemList)
                    {
                        prItem.Meterage = pItem.Meterage;
                        //costList.Add(prItem);
                        double pk = 0;
                        if (prItem.Volume == 0)
                        {
                            pk = prItem.Weight * pItem.Coverage;
                            AddToStage(stagestr, "Consumable", double.Parse(ConsAreatb.Text), 0, prItem.Weight, prItem.Meterage / pk,
                                            pItem.Meterage / pk * prItem.Weight,
                                            prItem.Unit, prItem.Name,
                                            prItem.Price,
                                            pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        }
                        else if (prItem.Weight == 0)
                        {
                            pk = prItem.Volume * pItem.Coverage;
                            AddToStage(stagestr, "Consumable", double.Parse(ConsAreatb.Text), 0, prItem.Volume, prItem.Meterage / pk,
                                            (pItem.Meterage / pk) * prItem.Volume,
                                            prItem.Unit,
                                            prItem.Name,
                                            prItem.Price,
                                            pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        }
                        else
                        {
                            pk = prItem.Volume * pItem.Coverage;
                            AddToStage(stagestr, "Consumable", double.Parse(ConsAreatb.Text), 0, prItem.Volume, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                        }
                    }
                }
            }
            else
            {
                foreach (Pack pItem in ConsumesList)
                {
                    foreach (Priceitem prItem in pItem.PriceItemList)
                    {
                        prItem.Meterage = pItem.Meterage;
                        //costList.Add(prItem);
                        double pk = 0;
                        if (prItem.Volume == 0)
                        {
                            pk = prItem.Weight * pItem.Coverage;
                            //AddToStage(stagestr, "Consumable", prItem.Meterage, 0, prItem.Weight * pk, prItem.Meterage * pk, pItem.Meterage * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price * pk, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                            AddToStage(stagestr, "Consumable", 0, double.Parse(ConsMultitb.Text), prItem.Volume, pItem.Meterage, pItem.Meterage * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage * prItem.Price, prItem.Name);
                        }
                        else if (prItem.Weight == 0)
                        {
                            pk = prItem.Volume * pItem.Coverage;
                            AddToStage(stagestr, "Consumable", double.Parse(ConsMultitb.Text), 0, prItem.Volume, pItem.Meterage, pItem.Meterage * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage * prItem.Price, prItem.Name);
                        }
                        else
                        {
                            pk = prItem.Volume * pItem.Coverage;
                            AddToStage(stagestr, "Consumable", double.Parse(ConsMultitb.Text), 0, prItem.Volume, pItem.Meterage, pItem.Meterage * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage * prItem.Price, prItem.Name);
                        }
                    }
                }
            }

            populateStage();
        }
        private void PrepaddtoStagebtn_Click(object sender, EventArgs e)
        {
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
            PrepaddtoStage();
        }

        private void ChevClearBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (ChevronLv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < ChevronLv.SelectedIndices.Count; i++)
                {
                    Selected = ChevronLv.SelectedIndices[0];
                    chevronList.RemoveAt(Selected);
                }
                Populate();
            }
        }

        private void StageFeditbtn_Click(object sender, EventArgs e)
        {
            newBackend.lists.COstageList = StageList;
            CostEdit costingEditor = new CostEdit(StageList, CostingList, newBackend);
            costingEditor.Owner = this;
            costingEditor.Show();
        }

        private void PrepAddbtn_Click(object sender, EventArgs e)
        {
            string preptype = "";
            int Selected = 0;
            this.errorProvider1.Clear();
            if (Prepsyslv.SelectedIndices.Count > 0)
            {
                Selected = Prepsyslv.SelectedIndices[0];
                string item = Prepsyslv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        preptype = item;
                        break;
                    }
                }
            }

            //string floortype = this.FloorType_Cb.Text;

            double Area = 0;

            bool isNumericA = double.TryParse(this.PrepAreatb.Text, out Area);
            if (!isNumericA)
            {
                this.errorProvider1.SetError(this.PrepAreatb, "Needs to be AlphaNumerical");
            }
            else
            {
                if (preptype != "" && isNumericA)
                {
                    foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                    {
                        if (psItem.BaseMat == preptype)
                        {
                            foreach (Pack pItem in psItem.PackList)
                            {
                                Pack newp = new Pack();
                                newp.Meterage = Area;
                                newp.Name = pItem.Name;
                                newp.Price = pItem.Price;
                                newp.Weight = pItem.Weight;
                                newp.Volume = pItem.Volume;
                                newp.Unit = pItem.Unit;
                                newp.Coverage = pItem.Coverage;
                                newp.ParentSystemName = psItem.BaseMat;
                                newp.PriceItemList = pItem.PriceItemList;
                                PrepList.Add(newp);
                            }
                        }
                    }
                }
                PopulatePrepList();
            }
        }

        private void Pareacb_CheckedChanged(object sender, EventArgs e)
        {
            if (Pareacb.Checked)
            {
                Parealmcb.Checked = false;
            }
        }

        private void Parealmcb_CheckedChanged(object sender, EventArgs e)
        {
            if (Parealmcb.Checked)
            {
                Pareacb.Checked = false;
            }
        }

        private void PrepRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (PrepListlv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < PrepListlv.SelectedIndices.Count; i++)
                {
                    Selected = PrepListlv.SelectedIndices[0];
                    PrepList.RemoveAt(Selected);
                }
                PopulatePrepList();
            }
        }

        private void Ncalc_Click(object sender, EventArgs e)
        {
            populateStage();
            cocalc();
            showEstTotal();
        }

        private void cocalc()
        {
            CostingList.Clear();
            double HoursCost = 0;
            foreach (CSstage Stageditem in StageList)
            {
                double packsAmnt = Stageditem.packsAmnt;
                double Ammount = Stageditem.Ammount;
                string symbol = Stageditem.symbol;
                string inputname = Stageditem.inputname;
                double cost = Stageditem.cost;
                double Packvol = Stageditem.Packvol;
                double price = Stageditem.price;

                bool wasinList = false;
                if (CombRescb.Checked == true)
                {
                    foreach (Costing Costitem in CostingList)
                    {
                        if (Costitem.inputname == inputname)
                        {
                            if (symbol == "Hour")
                            {
                                HoursCost += price;
                            }
                            else
                            {
                                Costitem.packsAmnt += packsAmnt;
                                Costitem.Ammount += Ammount;
                                Costitem.symbol = symbol;
                                Costitem.cost = cost;
                                Costitem.Packvol = Packvol;
                                Costitem.price += price;
                            }
                            wasinList = true;
                        }
                    }
                    if (!wasinList)
                    {

                        Costing nc = new Costing();
                        nc.packsAmnt = packsAmnt;
                        nc.Ammount = Ammount;
                        nc.symbol = symbol;
                        nc.inputname = inputname;
                        nc.cost = cost;
                        nc.Packvol = Packvol;
                        nc.price = price;
                        CostingList.Add(nc);
                    }
                }
                else
                {
                    Costing nc = new Costing();
                    nc.packsAmnt = packsAmnt;
                    nc.Ammount = Ammount;
                    nc.symbol = symbol;
                    nc.inputname = inputname;
                    nc.cost = cost;
                    nc.Packvol = Packvol;
                    nc.price = price;
                    CostingList.Add(nc);
                }
            }
            //Populate();
            //CalculateHours();
            //DisplaySummary();
            showEstTotal();
        }

        private void RCaddBtn_Click(object sender, EventArgs e)
        {
            string floortype = "";
            int Selected = 0;
            if (FloorRCLv.SelectedIndices.Count > 0)
            {
                Selected = FloorRCLv.SelectedIndices[0];
                string item = FloorRCLv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        floortype = item;
                        break;
                    }
                }
            }
            //string floortype = this.FloorType_Cb.Text;
            double Area = 0;
            bool isNumeric = double.TryParse(this.RCAreaTb.Text, out Area);
            if (floortype != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == floortype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Area;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            RCfloorList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }

        private void BCaddBtn_Click(object sender, EventArgs e)
        {
            string floortype = "";
            int Selected = 0;
            if (FloorBCFLv.SelectedIndices.Count > 0)
            {
                Selected = FloorBCFLv.SelectedIndices[0];
                string item = FloorBCFLv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        floortype = item;
                        break;
                    }
                }
            }
            //string floortype = this.FloorType_Cb.Text;
            double Area = 0;
            bool isNumeric = double.TryParse(this.BCFAreaTb.Text, out Area);
            if (floortype != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == floortype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Area;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            BCFfloorList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }

        private void ESLaddBtn_Click(object sender, EventArgs e)
        {
            string floortype = "";
            int Selected = 0;
            if (FloorESLLv.SelectedIndices.Count > 0)
            {
                Selected = FloorESLLv.SelectedIndices[0];
                string item = FloorESLLv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        floortype = item;
                        break;
                    }
                }
            }
            //string floortype = this.FloorType_Cb.Text;
            double Area = 0;
            bool isNumeric = double.TryParse(this.ESLAreaTb.Text, out Area);
            if (floortype != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == floortype)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Area;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            newp.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            ESLfloorList.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }

        private void RCRemBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (RCLv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < RCLv.SelectedIndices.Count; i++)
                {
                    Selected = RCLv.SelectedIndices[0];
                    RCfloorList.RemoveAt(Selected);
                }
                Populate();
            }
        }

        private void BCRemBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (BCFLv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < BCFLv.SelectedIndices.Count; i++)
                {
                    Selected = BCFLv.SelectedIndices[0];
                    BCFfloorList.RemoveAt(Selected);
                }
                Populate();
            }
        }

        private void ESLRemBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (ESLLv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < ESLLv.SelectedIndices.Count; i++)
                {
                    Selected = ESLLv.SelectedIndices[0];
                    ESLfloorList.RemoveAt(Selected);
                }
                Populate();
            }
        }

        private void RCclearBtn_Click(object sender, EventArgs e)
        {
            RCfloorList.Clear();
            //FloorType_Cb.Text = "";
            RCAreaTb.Text = "";
            Populate();
        }

        private void BCclearBtn_Click(object sender, EventArgs e)
        {
            BCFfloorList.Clear();
            //FloorType_Cb.Text = "";
            BCFAreaTb.Text = "";
            Populate();
        }

        private void ESLClearBtn_Click(object sender, EventArgs e)
        {
            ESLfloorList.Clear();
            //FloorType_Cb.Text = "";
            ESLAreaTb.Text = "";
            Populate();
        }

        private void RCAddItoSbtn_Click(object sender, EventArgs e)
        {
            string stagestr = RCStagecb.Text;
            string floorstr = "Floorsystem-RollCoat";
            Stagelv.Items.Clear();

            foreach (Pack pItem in RCfloorList)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(RCAreaTb.Text), 0, prItem.Weight, prItem.Meterage / pk,
                                        pItem.Meterage / pk * prItem.Weight,
                                        prItem.Unit, prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(RCAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk,
                                        (pItem.Meterage / pk) * prItem.Volume,
                                        prItem.Unit,
                                        prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(RCAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];

        }

        private void BCFAddItoSbtn_Click(object sender, EventArgs e)
        {
            string stagestr = BCFStagecb.Text;
            string floorstr = "Floorsystem-BroadCast-flake";
            Stagelv.Items.Clear();

            foreach (Pack pItem in BCFfloorList)
            {


                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(BCFAreaTb.Text), 0, prItem.Weight, prItem.Meterage / pk,
                                        pItem.Meterage / pk * prItem.Weight,
                                        prItem.Unit, prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(BCFAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk,
                                        (pItem.Meterage / pk) * prItem.Volume,
                                        prItem.Unit,
                                        prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(BCFAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void ESLAddItoSbtn_Click(object sender, EventArgs e)
        {
            string stagestr = ESLStagecb.Text;
            string floorstr = "Floorsystem-EpoxySL";
            Stagelv.Items.Clear();

            foreach (Pack pItem in ESLfloorList)
            {

                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(ESLAreaTb.Text), 0, prItem.Weight, prItem.Meterage / pk,
                                        pItem.Meterage / pk * prItem.Weight,
                                        prItem.Unit, prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(ESLAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk,
                                        (pItem.Meterage / pk) * prItem.Volume,
                                        prItem.Unit,
                                        prItem.Name,
                                        prItem.Price,
                                        pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, floorstr, double.Parse(ESLAreaTb.Text), 0, prItem.Volume, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void JointaddtoStagebtn_Click(object sender, EventArgs e)
        {
            string stagestr = JointStagecb.Text;
            Stagelv.Items.Clear();

            foreach (Pack pItem in Jointlist)
            {
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    prItem.Meterage = pItem.Meterage;
                    //costList.Add(prItem);
                    double pk = 0;
                    if (prItem.Volume == 0)
                    {
                        pk = prItem.Weight * pItem.Coverage;
                        AddToStage(stagestr, "Joints", 0, double.Parse(JointLMtb.Text), prItem.Weight, prItem.Meterage / pk, pItem.Meterage / pk * prItem.Weight, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else if (prItem.Weight == 0)
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Joints", 0, double.Parse(JointLMtb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                    else
                    {
                        pk = prItem.Volume * pItem.Coverage;
                        AddToStage(stagestr, "Joints", 0, double.Parse(JointLMtb.Text), prItem.Volume, pItem.Meterage / pk, pItem.Meterage / pk * prItem.Volume, prItem.Unit, prItem.Name, prItem.Price, pItem.Meterage / pk * prItem.Price, pItem.ParentSystemName);
                    }
                }
            }
            populateStage();
            //HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[0];
        }

        private void Jointaddbtn_Click(object sender, EventArgs e)
        {
            string Type = this.JointSyscb.Text;
            double Length;
            bool isNumeric = double.TryParse(this.JointLMtb.Text, out Length);
            if (Type != "" && isNumeric)
            {
                foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                {
                    if (psItem.BaseMat == Type)
                    {
                        foreach (Pack pItem in psItem.PackList)
                        {
                            Pack newp = new Pack();
                            newp.Meterage = Length;
                            newp.Name = pItem.Name;
                            newp.Price = pItem.Price;
                            newp.Weight = pItem.Weight;
                            newp.Volume = pItem.Volume;
                            newp.Unit = pItem.Unit;
                            newp.Coverage = pItem.Coverage;
                            pItem.ParentSystemName = psItem.BaseMat;
                            newp.PriceItemList = pItem.PriceItemList;
                            Jointlist.Add(newp);
                        }
                    }
                }
            }
            Populate();
        }

        private void JointRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (JointsLv.SelectedIndices.Count > 0)
            {
                Selected = JointsLv.SelectedIndices[0];
                Jointlist.RemoveAt(Selected);
                Populate();
            }
        }

        private void Jointclearbtn_Click(object sender, EventArgs e)
        {
            Jointlist.Clear();
            JointSyscb.Text = "";
            JointLMtb.Text = "";
            Populate();
        }

        private void Hoursimportbtn_Click(object sender, EventArgs e)
        {
            importToHrsList();
        }

        private void importToHrsList()
        {
            stagehrslist.Clear();
            foreach (CSstage Stgi in StageList)
            {
                bool inlist = false;
                //foreach (stagedhrs stritem in stagehrslist)
                //{
                //    string new1 = stritem.PackSystemName;
                //    string new2 = Stgi.PackSystemName;

                //    if (Stgi.Staging == stritem.Staging)
                //    {
                //        if (Stgi.Systems == stritem.Systems)
                //        {
                //            if (stritem.PackSystemName == Stgi.PackSystemName)
                //            {
                //                inlist = true;
                //            }
                //        }
                //    }
                //}
                if (!inlist)
                {
                    stagedhrs stghrsItem = new stagedhrs();
                    stghrsItem.Staging = Stgi.Staging;
                    stghrsItem.Systems = Stgi.Systems;
                    stghrsItem.Area = Stgi.Area;
                    stghrsItem.Lm = Stgi.Lm;
                    stghrsItem.PackSystemName = Stgi.PackSystemName + " " + Stgi.inputname;
                    stagehrslist.Add(stghrsItem);
                }
            }
            PopulateHrs();
        }

        private void PopulateHrs()
        {
            this.StageHourslv.Items.Clear();
            //stagehrslist.Sort((x, y) => x.Staging.CompareTo(y.Staging));
            foreach (stagedhrs stageitem in stagehrslist)
            {
                string[] newArr = new string[10];
                newArr[0] = stageitem.Staging;
                newArr[1] = stageitem.Systems;
                newArr[2] = stageitem.Area.ToString();
                newArr[3] = stageitem.Lm.ToString();
                newArr[4] = stageitem.PackSystemName;
                ListViewItem itm = new ListViewItem(newArr);
                this.StageHourslv.Items.Add(itm);
                
            }
        }

        private void Hrattachbtn_Click(object sender, EventArgs e)
        {
            if (Mnametb.Text != "")
            {
                stagedhrs newStgHr = new stagedhrs();
                Hrrates newHrItem = new Hrrates();
                string MethodName = "";
                int Selected1 = 0;
                bool ok1 = false;
                bool ok2 = false;
                if (StageHourslv.SelectedIndices.Count > 0)
                {
                    Selected1 = StageHourslv.SelectedIndices[0];
                    newStgHr = stagehrslist[Selected1];
                    ok1 = true;
                }

                if (Hoursweightlv.SelectedIndices.Count > 0)
                {
                    ListViewItem item = Hoursweightlv.SelectedItems[0];
                    MethodName = item.Text;
                    ok2 = true;
                }

                if (ok1 && ok2)
                {
                    foreach (Hrrates Hritem in newBackend.lists.HrrateList)
                    {
                        if (Hritem.hrsname == MethodName)
                        {
                            newHrItem = Hritem;
                        }
                    }

                    HrCombo newHrCitem = new HrCombo();
                    newHrCitem.Method = MethodName;
                    newHrCitem.SysName = Mnametb.Text;
                    newHrCitem.Stage = Hrsstagetb.Text;
                    newHrCitem.Rate = newHrItem.hrsWeight;
                    double Lmin = newStgHr.Lm;
                    double Areain = newStgHr.Area;
                    if (Lmin > 0)
                    {
                        newHrCitem.Length = newStgHr.Lm;
                        newHrCitem.Area = 0;
                        newHrCitem.Result = newHrCitem.Length * newHrCitem.Rate;
                    }
                    else if (Areain > 0)
                    {
                        newHrCitem.Area = newStgHr.Area;
                        newHrCitem.Length = 0;
                        newHrCitem.Result = newHrCitem.Area * newHrCitem.Rate;
                    }
                    else if (Areain > 0 && Lmin > 0)
                    {
                        newHrCitem.Area = newStgHr.Area;
                        newHrCitem.Length = newStgHr.Lm;
                        newHrCitem.Result = newStgHr.Lm * newHrCitem.Area * newHrCitem.Rate;
                    }
                    newBackend.lists.HrComboList.Add(newHrCitem);
                    updateHRScombo();
                }
            }
        }

        private void HrDeletebtn_Click(object sender, EventArgs e)
        {
            // HrCombolv
            int Selected = 0;
            if (HrCombolv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < HrCombolv.SelectedIndices.Count; i++)
                {
                    Selected = HrCombolv.SelectedIndices[0];
                    newBackend.lists.HrComboList.RemoveAt(Selected);

                }
                updateHRScombo();
            }
        }

        private void StageHourslv_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Mnametb.Text = "";
            int selected = 0;
            if (StageHourslv.SelectedIndices.Count > 0)
            {
                selected = StageHourslv.SelectedIndices[0];
                string output = stagehrslist[selected].PackSystemName;
                Hrsstagetb.Text = stagehrslist[selected].Staging;
                Mnametb.Text = output;
            }
        }

        private void Hrscalcbtn_Click(object sender, EventArgs e)
        {

            EstHrsDgv.Rows.Clear();
            Updateboxes();

            foreach (HrCombo hrcitem in newBackend.lists.HrComboList)
            {
                EstHrsDgv.Rows.Add(hrcitem.Area, hrcitem.Length, hrcitem.Stage, hrcitem.SysName, hrcitem.Method, "", hrcitem.Result);
            }
            HourGraphTabs.SelectedTab = HourGraphTabs.TabPages[1];

            //List<string> Localstageslist = new List<string>();

            //foreach (CSstage stgitem in StageList)
            //{
            //    string currentstring = stgitem.Staging;
            //    if (!Localstageslist.Contains(currentstring))
            //    {
            //        Localstageslist.Add(stgitem.Staging);
            //    }
            //}
            //Localstageslist.Sort();
            //foreach (string stritem in Localstageslist)
            //{
            //    SummaryLB.Items.Add("---------------- " + stritem + " ----------------");
            //    foreach (CSstage stgitem in StageList)
            //    {
            //        if (stgitem.Staging == stritem)
            //        {
            //            SummaryLB.Items.Add("--" + stgitem.inputname + " $" + stgitem.price / stgitem.Area + " per" + stgitem.symbol + " -- $" + stgitem.price);
            //        }
            //    }
            //}
        }
        public void Updateboxes()
        {
            foreach (Hrrates hritem in newBackend.lists.HrrateList)
            {
                if (hritem.hrsname == "HourMulti")
                {
                    MainHrRatetb.Text = hritem.hrsWeight.ToString();
                }
                if (hritem.hrsname == "Costmarkup")
                {
                    CostMarkuptb.Text = hritem.hrsWeight.ToString();
                }
            }
        }

        private void Esthrsprocbtn_Click(object sender, EventArgs e)
        {
            double count = EstHrsDgv.Rows.Count;
            double sum = 0;
            double sum2 = 0;
            HoursLb.Items.Clear();

            for (int i = StageList.Count - 1; i >= 0; i--)
            {
                if (StageList[i].symbol == "Hour")
                {
                    StageList.RemoveAt(i);
                }
            }
            string Area = "";//area
            string obj7 = "";//Length
            string obj5 = "";//stage
            string obj6 = "";//system
            string obj2 = "";//Name
            string obj3 = "";

            for (int i = 0; i < count - 1; i++)
            {
                Area = EstHrsDgv.Rows[i].Cells[0].Value.ToString();//area
                obj7 = EstHrsDgv.Rows[i].Cells[1].Value.ToString();//Length
                obj5 = EstHrsDgv.Rows[i].Cells[2].Value.ToString();//stage
                obj6 = EstHrsDgv.Rows[i].Cells[3].Value.ToString();//system
                obj2 = EstHrsDgv.Rows[i].Cells[4].Value.ToString();//Name
                obj3 = "";
                if (EstHrsDgv.Rows[i].Cells[5].Value != null)
                {
                    obj3 = EstHrsDgv.Rows[i].Cells[5].Value.ToString();//tech hrs
                    //if (obj3 == "")
                    //    obj3 = "0";
                }
                string obj4 = EstHrsDgv.Rows[i].Cells[6].Value.ToString();//calc hrs

                if (obj4 != "")
                {
                    sum2 += double.Parse(obj4);
                }
                if (obj3 != "")
                {
                    sum += double.Parse(obj3);
                }
                string output = Area + "m2 - " + " - " + obj2 + " - " + obj3 + " Hour";

                double Hrsss1 = 0;
                double Hrsss2 = 0;
                bool isNumeric1 = double.TryParse(obj3, out Hrsss1);
                bool isNumeric2 = double.TryParse(obj4, out Hrsss2);

                if (isNumeric1)
                {
                    AddToStage(obj5, "Hour", double.Parse(Area), 0, double.Parse(obj3), double.Parse(obj3), Hrsss1, "Hour", obj6 + " " + obj2, double.Parse(MainHrRatetb.Text), double.Parse(obj3) * double.Parse(MainHrRatetb.Text), obj6);
                }
                else
                {
                    if (isNumeric2)
                    {
                        AddToStage(obj5, "Hour", double.Parse(Area), 0, double.Parse(obj4), double.Parse(obj4), Hrsss2, "Hour", obj6 + " " + obj2, double.Parse(MainHrRatetb.Text), double.Parse(obj4) * double.Parse(MainHrRatetb.Text), obj6);
                    }
                }

                HoursLb.Items.Add(output);

            }
            TotalHrsTb.Text = sum2.ToString();
            if(sum > 0)
            {
                Additiontb.Text = (sum - sum2).ToString();
            }
                       
            Esthrstotaltb.Text = (double.Parse(TotalHrsTb.Text) + double.Parse(Additiontb.Text)).ToString();
            populateStage();

        }

        private void StageHrsclearbtn_Click(object sender, EventArgs e)
        {
            stagehrslist.Clear();  
            newBackend.lists.HrComboList.Clear();
            Hrsstagetb.Text = "";
            Mnametb.Text = "";
            HrCombolv.Items.Clear();
            PopulateHrs();
            updateHRScombo();
        }

        private void hrsclearbtn_Click(object sender, EventArgs e)
        {
            newBackend.lists.HrComboList.Clear();
            updateHRScombo();
        }

        private void CombRescb_CheckedChanged(object sender, EventArgs e)
        {
            Ncalc_Click(sender, e);
        }

        private void Addtoboxbtn_Click(object sender, EventArgs e)
        {
            Mnametb.Text = "";
            int selected = 0;
            if (StageHourslv.SelectedIndices.Count > 0)
            {
                selected = StageHourslv.SelectedIndices[0];
                string output = stagehrslist[selected].PackSystemName;
                Hrsstagetb.Text = stagehrslist[selected].Staging;
                Mnametb.Text = output;
            }
        }

        private void ConsAddbtn_Click(object sender, EventArgs e)
        {
            string Constype = "";
            int Selected = 0;
            if (Conslv.SelectedIndices.Count > 0)
            {
                Selected = Conslv.SelectedIndices[0];
                string item = Conslv.Items[Selected].Text;
                foreach (ProductSystem FlItem in newBackend.lists.SystemsList)
                {
                    if (item == FlItem.BaseMat)
                    {
                        //GcNcTb.Text = GCmpitm;
                        //GCommentList.Remove(GCmpitm);
                        //Populate();
                        Constype = item;
                        break;
                    }
                }
            }
            //string floortype = this.FloorType_Cb.Text;
            double multi = 0;
            bool isNumericM = double.TryParse(this.ConsMultitb.Text, out multi);
            double Area = 0;
            bool isNumericA = double.TryParse(this.ConsAreatb.Text, out Area);
            bool pass1 = true;
            bool pass2 = true;

            if (!isNumericM && Areamulticb.Checked)
            {
                this.errorProvider1.Clear();
                //this.errorProvider1.SetError(this.ConsMultitb, "Needs to be AlphaNumerical");
                pass1 = false;

            }

            if (!isNumericA && !Areamulticb.Checked)
            {
                this.errorProvider1.Clear();
                //this.errorProvider1.SetError(this.ConsAreatb, "Needs to be AlphaNumerical");
                pass2 = false;
            }

            if (pass2)
            {
                if (Constype != "" && isNumericA)
                {
                    foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                    {
                        if (psItem.BaseMat == Constype)
                        {
                            foreach (Pack pItem in psItem.PackList)
                            {
                                Pack newp = new Pack();
                                newp.Meterage = Area;
                                newp.Name = pItem.Name;
                                newp.Price = pItem.Price;
                                newp.Weight = pItem.Weight;
                                newp.Volume = pItem.Volume;
                                newp.Unit = pItem.Unit;
                                newp.Coverage = pItem.Coverage;
                                newp.ParentSystemName = psItem.BaseMat;
                                newp.PriceItemList = pItem.PriceItemList;
                                ConsumesList.Add(newp);
                            }
                        }
                    }
                    PopulateConsList();
                }
            }

            if (pass1)
            {
                if (Constype != "" && isNumericM)
                {
                    foreach (ProductSystem psItem in newBackend.lists.SystemsList)
                    {
                        if (psItem.BaseMat == Constype)
                        {
                            foreach (Pack pItem in psItem.PackList)
                            {
                                bool wasinList = false;
                                foreach (Pack Cspack in ConsumesList)
                                {
                                    if (Cspack.Name == pItem.Name)
                                    {
                                        wasinList = true;
                                        Cspack.Name = pItem.Name;
                                        Cspack.Unit = pItem.Unit;
                                        Cspack.Coverage += pItem.Coverage * multi;
                                        Cspack.Meterage += multi;
                                        Cspack.Price += pItem.Price * multi;
                                        Cspack.Weight += pItem.Weight * multi;
                                        Cspack.Volume += pItem.Volume * multi;
                                        Cspack.ParentSystemName = psItem.BaseMat;
                                        Cspack.PriceItemList = pItem.PriceItemList;

                                    }
                                }
                                if (!wasinList)
                                {
                                    Pack newp = new Pack();
                                    newp.Name = pItem.Name;
                                    newp.Meterage = multi;
                                    newp.Price = pItem.Price * multi;
                                    newp.Weight = pItem.Weight * multi;
                                    newp.Volume = pItem.Volume * multi;
                                    newp.Unit = pItem.Unit;
                                    newp.Coverage = pItem.Coverage * multi;
                                    newp.ParentSystemName = psItem.BaseMat;
                                    newp.PriceItemList = pItem.PriceItemList;
                                    ConsumesList.Add(newp);
                                }
                            }
                        }
                    }
                }
                PopulateConsList();
            }

            //PrepaddtoStage();
        }

        private void ConsAddtostagebtn_Click(object sender, EventArgs e)
        {
            ConsaddtoStage();
        }

        private void ConsClearbtn_Click(object sender, EventArgs e)
        {
            ConsumesList.Clear();
            PopulateConsList();
        }

        private void PrepClearbtn_Click(object sender, EventArgs e)
        {
            PrepList.Clear();
            PopulatePrepList();
        }

        private void ConsumesSearchtb_TextChanged(object sender, EventArgs e)
        {
            UpdateConsumesBox();
        }

        private void clearListsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StageClearbtn_Click(sender, e);
        }

        private void clearStageBoxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.PUCStagecb.Text = "";
            this.CovStagecb.Text = "";
            this.sfStagecb.Text = "";
            this.lmStagecb.Text = "";
            this.chvStagecb.Text = "";
            this.pcStagecb.Text = "";
            this.PrepStagecb.Text = "";
            this.RCStagecb.Text = "";
            this.BCFStagecb.Text = "";
            this.ESLStagecb.Text = "";
            this.JointStagecb.Text = "";
            this.ConsAtoScb.Text = "";

            this.PUCStagecb.Items.Clear();
            this.CovStagecb.Items.Clear();
            this.sfStagecb.Items.Clear();
            this.lmStagecb.Items.Clear();
            this.chvStagecb.Items.Clear();
            this.pcStagecb.Items.Clear();
            this.PrepStagecb.Items.Clear();
            this.RCStagecb.Items.Clear();
            this.BCFStagecb.Items.Clear();
            this.ESLStagecb.Items.Clear();
            this.JointStagecb.Items.Clear();
            this.ConsAtoScb.Items.Clear();
        }

        private void EsthrsRemBtn_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell oneCell in EstHrsDgv.SelectedCells)
            {
                if (oneCell.Selected)
                {
                    if (oneCell.RowIndex != EstHrsDgv.Rows.Count - 1)
                    {
                        EstHrsDgv.Rows.RemoveAt(oneCell.RowIndex);
                    }
                }
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }
        public void ForceClose()
        {
            MessageBox.Show("Your Connection to the Server has been disconnect", "Message,", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Close();
        }
        private void forceRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            newBackend.lists.RelinkPackLinks();
            newBackend.lists.RelinkSystemDataLinks();
            Populate();
            UpdateFloorsysBox();
            updateQuotebox();
        }

        private void Areamulticb_CheckedChanged(object sender, EventArgs e)
        {
            ConsumesList.Clear();
            PopulateConsList();

            if (Areamulticb.Checked)
            {
                ConsAreatb.Enabled = true;
                ConsMultitb.Enabled = false;
                ConsMultitb.Text = "";
                ConsAreatb.Text = "0";
            }
            else
            {
                ConsAreatb.Text = "";
                ConsAreatb.Enabled = false;
                ConsMultitb.Enabled = true;
                ConsMultitb.Text = "1";
            }
        }

        private void Addselbtn_Click(object sender, EventArgs e)
        {
            string stritem = "";
            int Selected = 0;
            if (Additionslv.SelectedIndices.Count > 0)
            {
                Selected = Additionslv.SelectedIndices[0];
                string item = Additionslv.Items[Selected].Text;
                foreach (string strItem in newBackend.lists.AdditionsList)
                {
                    if (item == strItem)
                    {
                        stritem = item;
                        break;
                    }
                }
                Additemtb.Text = stritem;
            }
        }



        private void Addsaddbtn_Click(object sender, EventArgs e)
        {
            string stritem =  Additemtb.Text;
            string strdesc = Adddesctb.Text;
            string stramnt = Addsammttb.Text;
            string strcost = Addscosttb.Text;

            if(stritem != "" && strcost != "")
            {
                Adds NewAdditional = new Adds();
                NewAdditional.ItemName = stritem;
                NewAdditional.Desc = strdesc;
                NewAdditional.Ammount = stramnt;
                NewAdditional.Cost = strcost;
                newBackend.lists.AddsList.Add(NewAdditional);
            }
            populateAdds();
            cocalc();
            Additemtb.Text = "";
            Adddesctb.Text = "";
            Addsammttb.Text = "";
            Addscosttb.Text = "";
        }

        private void Addrembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Addslv.SelectedIndices.Count > 0)
            {
                for (int i = 0; i < Addslv.SelectedIndices.Count; i++)
                {
                    Selected = Addslv.SelectedIndices[0];
                    newBackend.lists.AddsList.RemoveAt(Selected);

                }
                populateAdds();
            }
        }

        private void AddsClearbtn_Click(object sender, EventArgs e)
        {
            newBackend.lists.AddsList.Clear();
            populateAdds();
        }

        private void QuoteDispBtn_Click(object sender, EventArgs e)
        {
            double FinCost = 0;
            bool isNumeric = double.TryParse(finalCosttb.Text, out FinCost);
            //finalCosttb.Text
            if (isNumeric)
            {
                Quote QuoteForm = new Quote(newBackend, FinCost, QuotesTextbox.Text);
                QuoteForm.Show();
            }
            else
            {
                Quote QuoteForm = new Quote(newBackend, double.Parse(totalCosttb.Text), QuotesTextbox.Text);
                QuoteForm.Show();
            }
        }

        private void PurchaseOrderBtn_Click(object sender, EventArgs e)
        {
            PurchaseOrder PoExporter = new PurchaseOrder(newBackend);
            //PoExporter.Owner = this;
            PoExporter.Show();
        }
    }


    public class LveChartApp
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public double Value { get; set; }
    }

    public class DisplaySystemCosting
    {
        public string Name { get; set; }
        public double PriceValue { get; set; }
        public double Area { get; set; }
        public double StdhrRate { get; set; }

    }
    public class SummaryOutput
    {
        public string F1 { get; set; }
        public string F2 { get; set; }
        public string F3 { get; set; }

    }


}
