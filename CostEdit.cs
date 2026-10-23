using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Qepoxy
{
    public partial class CostEdit : Form
    {
        public List<Costing> LocalCostList = new List<Costing>();
        public List<Priceitem> LocalPriceList = new List<Priceitem>();
        public List<Priceitem> SearchPriceList = new List<Priceitem>();
        public List<CSstage> LocalStageList = new List<CSstage>();

        Backend newBe;        
        public CostEdit(List<CSstage> stageList,List<Costing> costList,Backend newBackend)
        {
            InitializeComponent();
            newBe = newBackend;
            this.CostListlv.Scrollable = true;
            this.CostListlv.View = View.Details;
            this.CostListlv.GridLines = true;
            this.CostListlv.FullRowSelect = true;
            this.CostListlv.Columns.Add("Ammount", 120);
            this.CostListlv.Columns.Add("#", 50);
            this.CostListlv.Columns.Add("Units", 50);
            this.CostListlv.Columns.Add("Product", 160);
            this.CostListlv.Columns.Add("$ per Ammount", 80);
            this.CostListlv.Columns.Add("Est Price", 80);
            this.CostListlv.MultiSelect = true;
            //this.CostListlv.setItemsCanFocus(false);

            this.costStageLv.Scrollable = true;
            this.costStageLv.View = View.Details;
            this.costStageLv.GridLines = true;
            this.costStageLv.FullRowSelect = true;
            this.costStageLv.Columns.Add("Stage", 120);
            this.costStageLv.Columns.Add("System Type", 110);
            this.costStageLv.Columns.Add("#Qty", 70);
            this.costStageLv.Columns.Add("Product", 160);
            this.costStageLv.Columns.Add("Kits", 120);
            this.costStageLv.Columns.Add("#", 50);
            this.costStageLv.Columns.Add("Units", 50);
            this.costStageLv.Columns.Add("$ per", 80);
            this.costStageLv.Columns.Add("Est Total", 80);


            this.Prodlv.Scrollable = true;
            this.Prodlv.View = View.Details;
            this.Prodlv.GridLines = true;
            this.Prodlv.FullRowSelect = true;
            this.Prodlv.Columns.Add("Product", 220);
            this.Prodlv.Columns.Add("Price", 50);
            this.Prodlv.Columns.Add("Weight", 40);
            this.Prodlv.Columns.Add("Volume", 40);
            this.Prodlv.Columns.Add("Coverage", 40);
            this.Prodlv.Columns.Add("Unit", 40);

            this.Unitcb.Items.AddRange(new object[] { "L" });
            this.Unitcb.Items.AddRange(new object[] { "Kg" });
            this.Unitcb.Items.AddRange(new object[] { "m" });
            this.Unitcb.Items.AddRange(new object[] { "Packets" });
            this.Unitcb.Items.AddRange(new object[] { "Kits" });
            this.Unitcb.Items.AddRange(new object[] { "Hours" });
            this.Unitcb.Items.AddRange(new object[] { "Rolls" });
            this.Unitcb.Items.AddRange(new object[] { "Items" });

            Addbtn.FlatStyle = FlatStyle.Flat;
            Addbtn.FlatAppearance.BorderColor = Color.Green;
            Addbtn.FlatAppearance.BorderSize = 1;
            
            Rembtn.FlatStyle = FlatStyle.Flat;
            Rembtn.FlatAppearance.BorderColor = Color.Red;
            Rembtn.FlatAppearance.BorderSize = 1;

            if (stageList.Count > 0)
            {
                foreach (CSstage stItem in stageList)
                {
                    LocalStageList.Add(stItem);
                }
            }

            if (costList.Count > 0)
            {
                foreach (Costing cItem in costList)
                {
                    LocalCostList.Add(cItem);
                }               
            }
            if (newBe.lists.PriceList.Count > 0)
            {
                foreach (Priceitem PrItem in newBe.lists.PriceList)
                {
                    LocalPriceList.Add(PrItem);
                    SearchPriceList.Add(PrItem);
                }
            }
            populateStage();
            Populate();
            this.FormClosing += Form_FormClosing;
        }
        private void Form_FormClosing(Object sender, FormClosingEventArgs e)
        {
            Main form1 = (Main)this.Owner;
            form1.CostingList = LocalCostList;
            form1.StageList = LocalStageList;
            form1.populateStage();
            form1.showEstTotal();
        }
        private void populateStage()
        {
            CalculateClist();
            this.costStageLv.Items.Clear();
            //LocalStageList.Sort((x, y) => x.Staging.CompareTo(y.Staging));
            foreach (CSstage stageitem in LocalStageList)
            {
                string[] newArr = new string[10];
                newArr[0] = stageitem.Staging;
                newArr[1] = stageitem.Systems;
                newArr[2] = stageitem.Area.ToString();
                // newArr[2] = stageitem.inputname;                
                double cstamnt = Math.Round(stageitem.packsAmnt, 3);
                newArr[3] = stageitem.inputname;
                newArr[4] = cstamnt.ToString();// + " kits";
                newArr[5] = stageitem.Ammount.ToString();
                newArr[6] = stageitem.symbol;
                newArr[7] = "$" + Math.Round(stageitem.cost, 2).ToString();
                newArr[8] = "$" + Math.Round(stageitem.price, 2).ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.costStageLv.Items.Add(itm);
            }
            Populate();
        }
        private void smallPopulateStage()
        {
            CalculateClist();
            this.costStageLv.Items.Clear();
            //LocalStageList.Sort((x, y) => x.Staging.CompareTo(y.Staging));
            foreach (CSstage stageitem in LocalStageList)
            {
                string[] newArr = new string[10];
                newArr[0] = stageitem.Staging;
                newArr[1] = stageitem.Systems;
                newArr[2] = stageitem.Area.ToString();
                // newArr[2] = stageitem.inputname;                
                double cstamnt = Math.Round(stageitem.packsAmnt, 3);
                newArr[3] = stageitem.inputname;
                newArr[4] = cstamnt.ToString();// + " kits";
                newArr[5] = stageitem.Ammount.ToString();
                newArr[6] = stageitem.symbol;
                newArr[7] = "$" + Math.Round(stageitem.cost, 2).ToString();
                newArr[8] = "$" + Math.Round(stageitem.price, 2).ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.costStageLv.Items.Add(itm);
            }            
        }
        public void Populate()
        {            
            CostListlv.Items.Clear();
            foreach (Costing Costitem in LocalCostList)
            {
                string[] newArr = new string[12];
                double cstamnt = Math.Round(Costitem.packsAmnt, 1);
                newArr[0] = cstamnt.ToString() + " kits";
                newArr[1] = Costitem.Ammount.ToString();
                newArr[2] = Costitem.symbol;
                newArr[3] = Costitem.inputname;
                newArr[4] = "$" + Math.Round(Costitem.cost,2).ToString();
                newArr[5] = "$" + Math.Round(Costitem.price,2).ToString();
                ListViewItem itm = new ListViewItem(newArr);
                this.CostListlv.Items.Add(itm);
            }
            double Markupc = 0;
            foreach (Hrrates hritem in newBe.lists.HrrateList)
            {
                if (hritem.hrsname == "Costmarkup")
                {
                    Markuptb.Text = hritem.hrsWeight.ToString();
                    Markupc = double.Parse("1." + Markuptb.Text);
                }
            }
            
            double totalcost = new double();
            totalcost = 0;
            foreach (Costing Costitem in LocalCostList)
            {
                Costitem.packsAmnt = Math.Ceiling(Costitem.packsAmnt);
                Costitem.Ammount = Costitem.Ammount;
                Costitem.symbol = Costitem.symbol;
                Costitem.cost = Costitem.cost;
                totalcost += Costitem.packsAmnt * Costitem.cost;
                
            }
            totalcost = Markupc * totalcost;
            EstimTb.Text = totalcost.ToString();

            this.Prodlv.Items.Clear();
            foreach (Priceitem Pitem in SearchPriceList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Price.ToString();
                newArr[2] = Pitem.Weight.ToString();
                newArr[3] = Pitem.Volume.ToString();
                newArr[4] = Pitem.Coverage.ToString();
                newArr[5] = Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.Prodlv.Items.Add(itm);
            }
        }
        private void clearErrorCheck()
        {
            errorProvider1.Clear();
        }

        private bool ErrorChecktb()
        {
            clearErrorCheck();
            bool output = false;
            double rst = 0;
            bool Qtb = double.TryParse(Quantitytb.Text, out rst);
            if (Quantitytb.Text == "" || !Qtb)
            {
                errorProvider1.SetError(Quantitytb, "Needs a value or Incorrect Input");
                output = true;
            }

            if (Unitcb.Text == "")
            {
                errorProvider1.SetError(Unitcb, "Needs a value or Incorrect Input");
                output = true;
            }

            if (PLstagetb.Text == "")
            {
                errorProvider1.SetError(PLstagetb, "Needs a value or Incorrect Input");
                output = true;
            }

            if (Plsystb.Text == "")
            {
                errorProvider1.SetError(Plsystb, "Needs a value or Incorrect Input");
                output = true;
            }
            return output;
        }
        //add button
        private void button4_Click(object sender, EventArgs e)
        {
           if(!ErrorChecktb())
            {
                double Qnty = 0;
                bool Qntypassed = double.TryParse(Quantitytb.Text, out Qnty);
                if (Qntypassed && Unitcb.Text != "")
                {
                    int Selected = 0;
                    if (Prodlv.SelectedIndices.Count > 0)
                    {
                        Selected = Prodlv.SelectedIndices[0];
                        string item = Prodlv.Items[Selected].Text;
                        foreach (Priceitem PItem in newBe.lists.PriceList)
                        {
                            if (PItem.Name == item)
                            {
                                if (Unitcb.Text == "Kits")
                                {
                                    double ratio = 0;
                                    CSstage newCstitm = new CSstage();
                                    newCstitm.inputname = PItem.Name;
                                    if (PItem.Volume == 0)
                                    {
                                        ratio = Qnty * PItem.Weight;
                                        newCstitm.packsAmnt = Qnty;
                                        newCstitm.Packvol = ratio;
                                        newCstitm.symbol = PItem.Unit;
                                        newCstitm.Ammount = ratio;
                                        newCstitm.PackSystemName = "none";
                                    }
                                    else if (PItem.Weight == 0)
                                    {
                                        ratio = Qnty * PItem.Volume;
                                        newCstitm.packsAmnt = Qnty;
                                        newCstitm.Packvol = ratio;
                                        newCstitm.symbol = PItem.Unit;
                                        newCstitm.Ammount = ratio;
                                        newCstitm.PackSystemName = "none";
                                    }
                                    else
                                    {
                                        ratio = Qnty * PItem.Volume;
                                        newCstitm.packsAmnt = Qnty;
                                        newCstitm.Packvol = ratio;
                                        newCstitm.symbol = PItem.Unit;
                                        newCstitm.Ammount = ratio;
                                        newCstitm.PackSystemName = "none";
                                    }
                                    newCstitm.cost = PItem.Price;
                                    newCstitm.price = Qnty * PItem.Price;
                                    newCstitm.Staging = PLstagetb.Text;
                                    newCstitm.PackSystemName = Plsystb.Text;
                                    LocalStageList.Add(newCstitm);
                                }
                                else
                                {
                                    double ratio = 0;
                                    CSstage newCstitm = new CSstage();
                                    newCstitm.inputname = PItem.Name;
                                    if (PItem.Volume == 0)
                                    {
                                        ratio = Qnty / PItem.Weight;
                                        newCstitm.packsAmnt = ratio;
                                        newCstitm.Ammount = PItem.Weight * ratio;
                                        newCstitm.symbol = Unitcb.Text;
                                        newCstitm.Packvol = PItem.Volume;
                                        newCstitm.PackSystemName = "none";

                                    }
                                    else if (PItem.Weight == 0)
                                    {
                                        ratio = Qnty / PItem.Volume;
                                        newCstitm.packsAmnt = ratio;
                                        newCstitm.Packvol = PItem.Volume;
                                        newCstitm.symbol = Unitcb.Text;
                                        newCstitm.Ammount = ratio * PItem.Volume;
                                        newCstitm.PackSystemName = "none";
                                    }
                                    else
                                    {
                                        ratio = Qnty / PItem.Volume;
                                        newCstitm.packsAmnt = ratio;
                                        newCstitm.Packvol = PItem.Volume;
                                        newCstitm.symbol = Unitcb.Text;
                                        newCstitm.Ammount = ratio * PItem.Volume;
                                        newCstitm.PackSystemName = "none";
                                    }
                                    newCstitm.cost = PItem.Price;
                                    newCstitm.price = ratio * PItem.Price;
                                    newCstitm.Staging = PLstagetb.Text;
                                    newCstitm.Systems = Plsystb.Text;
                                    LocalStageList.Add(newCstitm);
                                }
                            }                            
                        }                        
                    }
                    populateStage();
                }
            }           
        }

        private void Rembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            CSstage CItemEdit = new CSstage();
            if (costStageLv.SelectedIndices.Count > 0)
            {
                Selected = costStageLv.SelectedIndices[0];
                LocalStageList.RemoveAt(Selected);                
            }
            populateStage();
        }

        private void ExportPlbtn_Click(object sender, EventArgs e)
        {            
            if (LocalStageList.Count > 0)
            {
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                object start_of_doc = "\\startofdoc";

                List<CSstage> cstgList = new List<CSstage>();
                cstgList.Clear();
                List<string> Lslist = new List<string>();
                List<string> LsAreaList = new List<string>();
                foreach (CSstage stgitem in LocalStageList)
                {
                    cstgList.Add(stgitem);
                    string currentstring = stgitem.Staging;
                    if (!Lslist.Contains(currentstring))
                    {
                        Lslist.Add(stgitem.Staging);
                        if (stgitem.Area > 0 && stgitem.Lm <= 0)
                        {
                            LsAreaList.Add(stgitem.Area.ToString() + " m2");
                        }
                        else if(stgitem.Lm > 0 && stgitem.Area <= 0)
                        {
                            LsAreaList.Add(stgitem.Lm.ToString() + " Lm");
                                                    
                        }
                        else if(stgitem.Area > 0 && stgitem.Lm > 0)
                        {
                            if (stgitem.Systems == "Subfill")
                            {
                                LsAreaList.Add(stgitem.Area.ToString() + " m2, " + stgitem.Lm.ToString() + " mm");
                            }
                            else
                            {
                                LsAreaList.Add(stgitem.Area.ToString() + " m2, " + stgitem.Lm.ToString() + " Lm");
                            }
                        }                        
                    }
                }

                string[,] CstListArray;
                CstListArray = new string[LocalStageList.Count +2+ Lslist.Count, 8];
                int i = 1;
                int j = 0;
                foreach (string stritem in Lslist)
                {
                    //SummaryLB.Items.Add("---------------- " + stritem + " ----------------");
                    i++;
                    CstListArray[i, 1] = stritem;
                    CstListArray[i, 2] = LsAreaList[j];
                    CstListArray[i, 3] = "";
                    CstListArray[i, 4] = "";
                    CstListArray[i, 5] = "";
                    CstListArray[i, 6] = "";
                    CstListArray[i, 7] = "";
                    //CstListArray[i, 8] = "";
                    //CstListArray[i, 7] = "";
                    j++;
                    CSstage newcst = new CSstage();
                    newcst.Staging = stritem;
                    newcst.Systems = "";
                    newcst.inputname = "";
                    cstgList.Add(newcst);
                    foreach (CSstage Cstgitem in LocalStageList)
                    {
                        if(stritem == Cstgitem.Staging)
                        {
                            i++;
                            cstgList.Add(Cstgitem);
                            CstListArray[i, 1] = "";
                            CstListArray[i, 2] = Cstgitem.inputname;
                            CstListArray[i, 3] = Math.Round(Cstgitem.Ammount, 2).ToString() + " " + Cstgitem.symbol;
                            CstListArray[i, 4] = Cstgitem.Packvol.ToString() + " " + Cstgitem.symbol;
                            CstListArray[i, 5] = Math.Ceiling(Cstgitem.packsAmnt).ToString();
                            CstListArray[i, 6] = "□";
                            CstListArray[i, 7] = "";
                            //CstListArray[i, 8] = "";
                            //CstListArray[i, 3] = "$" + Math.Round(Coitem.cost, 2).ToString();
                            //CstListArray[i, 4] = "$" + Math.Round(Coitem.price, 2).ToString();
                            
                        }
                    }
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
                string path = wanted_path + "..\\..\\Data\\Qepoxy1.jpg";
                range.Application.ActiveDocument.InlineShapes.AddPicture(@path, ref oLinkToFile, ref oSaveWithDocument, ref oMissed);

               

                int spacing = 12;
                //Word.Paragraph oPara3z;
                //oPara3z = oDoc.Content.Paragraphs.Add(ref oMissing);
                //oPara3z.Format.SpaceAfter = spacing;
                //oPara3z.Range.InsertParagraphAfter();

                Word.Paragraph oPara1a;
                oPara1a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara1a.Range.Text = "Job Product Sheet";
                oPara1a.Range.Font.Bold = 1;
                oPara1a.Range.Font.Size = 12;
                oPara1a.Range.Font.Name = "calibri";
                oPara1a.Format.SpaceAfter = spacing;
                oPara1a.Range.InsertParagraphAfter();

                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3a.Range.Text = "Name: " + JnameTb.Text;
                oPara3a.Range.Font.Bold = 1;
                oPara3a.Range.Font.Size = 12;
                oPara3a.Range.Font.Name = "calibri";
                oPara3a.Format.SpaceAfter = spacing;
                oPara3a.Range.InsertParagraphAfter();

                Word.Paragraph oPara3c;
                oPara3c = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3c.Range.Text = "Quote #: " + Jobnumtb.Text;
                oPara3c.Range.Font.Bold = 1;
                oPara3c.Range.Font.Size = 12;
                oPara3c.Range.Font.Name = "calibri";
                oPara3c.Format.SpaceAfter = spacing;
                oPara3c.Range.InsertParagraphAfter();

                Word.Paragraph oPara3b;
                oPara3b = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3b.Range.Text = "Size: " + JsizeTb.Text;
                oPara3b.Range.Font.Bold = 1;
                oPara3b.Range.Font.Size = 12;
                oPara3b.Range.Font.Name = "calibri";
                oPara3b.Format.SpaceAfter = spacing;
                oPara3b.Range.InsertParagraphAfter();

                Word.Paragraph oPara2;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
                //int tablerows = cstgList.Count- Lslist.Count;
                int tablerows = i;
                Word.Table firstTable = oDoc.Tables.Add(oPara2.Range, tablerows, 7, ref oMissing, ref oMissing);
                int Basefontsize = 10;
                firstTable.Borders.Enable = 0;
                //float width = 20;
                firstTable.Columns[2].SetWidth(160,Word.WdRulerStyle.wdAdjustSameWidth);
                firstTable.Columns[3].SetWidth(50, Word.WdRulerStyle.wdAdjustSameWidth);
                firstTable.Rows.SetHeight(8, Word.WdRowHeightRule.wdRowHeightAuto);
                //firstTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[2].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[3].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[4].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[5].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[6].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[7].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
               // firstTable.Columns[8].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                foreach (Word.Row row in firstTable.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        if(cell.ColumnIndex == 1)
                        {
                            
                        }                        
                        //Header row  
                        if (cell.RowIndex == 1 && cell.ColumnIndex == 1)
                        {                            
                            cell.Range.Text = "Stage ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize+2;
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
                            cell.Range.Text = "Description: ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;                            

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 3)
                        {
                            
                            cell.Range.Text = "Qty:";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
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
                            
                            cell.Range.Text = "Size:";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 5)
                        {
                            cell.Range.Text = "#:";
                                                      
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 6)
                        {
                            cell.Range.Text = "Check #:";                            
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 7)
                        {
                            cell.Range.Text = "# Used:";

                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
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
                            if (cell.ColumnIndex == 6)
                            {                                     
                                //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                firstTable.Cell(cell.RowIndex, cell.ColumnIndex).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = 12;
                                cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;                            
                                cell.Range.Text = CstListArray[cell.RowIndex, cell.ColumnIndex];
                                
                            }
                            else
                            {
                                if(CstListArray[cell.RowIndex, cell.ColumnIndex].Contains("Hours"))
                                {

                                }
                                else
                                {
                                    cell.Range.Font.Bold = 0;
                                    //firstTable.Cell(cell.RowIndex, cell.ColumnIndex).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                                    //other format properties goes here  
                                    cell.Range.Font.Name = "calibri";
                                    cell.Range.Font.Size = Basefontsize;
                                    //cell.Range.Paragraphs.SpaceBefore = 0;
                                    cell.Range.Paragraphs.SpaceAfter = 0;
                                    cell.Range.Text = CstListArray[cell.RowIndex, cell.ColumnIndex];
                                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                }                               
                            }    
                        }
                        else
                        {
                            //cell.Range.Text = "$" + (Subtotal * 1.1).ToString();
                            //cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize;
                            //cell.Range.Paragraphs.SpaceBefore = 0;
                            //cell.Range.Paragraphs.SpaceAfter = 0;                            
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            //cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                }
                foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
                {
                    Word.Range header = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    Word.Paragraph h1 = header.Paragraphs.Add(header);
                    header.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    header.Text = "Version: 2.5";
                    header.Font.Size = 8;
                    header.Font.Bold = 0;
                    // header.InsertParagraphAfter();

                    Word.Range footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    Word.Paragraph p1 = footer.Paragraphs.Add(footer);
                    footer.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //footer.Text = "Version: 2.5";
                    footer.Font.Size = 8;
                    footer.Font.Bold = 0;
                    Word.Paragraph p3 = footer.Paragraphs.Add(footer);
                    p3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    p3.Range.Fields.Add(footer, Word.WdFieldType.wdFieldPage);
                }
            }            
        }

        private void SearchTb_TextChanged(object sender, EventArgs e)
        {
            Prodlv.Items.Clear();
            SearchPriceList.Clear();
            //StringComparison comp = StringComparison.CurrentCultureIgnoreCase;
            if (SearchTb.Text != "")
            {
                foreach (Priceitem PrItem in LocalPriceList)
                {
                    //LocalCostList.Add(cItem);                    
                    if (PrItem.Name.ToUpper().Contains(SearchTb.Text.ToUpper()))
                    {
                        SearchPriceList.Add(PrItem);
                    }
                }
            }
            else
            {
                foreach (Priceitem PrItem in LocalPriceList)
                {
                    SearchPriceList.Add(PrItem);
                }   
            }
            Populate();
        }

        private void upbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (costStageLv.SelectedIndices.Count > 0)
            {

                //Mnametb.Text = "";
                //int selected = 0;
                //if (StageHourslv.SelectedIndices.Count > 0)
                //{
                //    selected = StageHourslv.SelectedIndices[0];
                //    string output = stagehrslist[selected].Staging + " " + stagehrslist[selected].PackSystemName;
                //    Mnametb.Text = output;
                //}
                
                Selected = costStageLv.SelectedIndices[0];
                CSstage stritem = LocalStageList[Selected];
                
                foreach (CSstage Stgitm in LocalStageList)
                {
                    if (Stgitm.inputname == stritem.inputname)
                    {
                        
                        int index = LocalStageList.IndexOf(stritem);
                        if (index <= 0)
                        {  costStageLv.Items[index].Selected = true;  }
                        else
                        {
                            LocalStageList.RemoveAt(index);
                            LocalStageList.Insert(index - 1, stritem);
                            smallPopulateStage();
                            costStageLv.Items[index - 1].Selected = true;
                            break;
                        }
                    }
                }
            }
        }

        private void downbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (costStageLv.SelectedIndices.Count > 0)
            {

                //Mnametb.Text = "";
                //int selected = 0;
                //if (StageHourslv.SelectedIndices.Count > 0)
                //{
                //    selected = StageHourslv.SelectedIndices[0];
                //    string output = stagehrslist[selected].Staging + " " + stagehrslist[selected].PackSystemName;
                //    Mnametb.Text = output;
                //}

                Selected = costStageLv.SelectedIndices[0];
                CSstage stritem = LocalStageList[Selected];

                foreach (CSstage Stgitm in LocalStageList)
                {
                    if (Stgitm.inputname == stritem.inputname)
                    {
                        int index = LocalStageList.IndexOf(stritem);
                        if ((index+1) == costStageLv.Items.Count)
                        {
                            costStageLv.Items[index].Selected = true;
                        }
                        else
                        {
                            LocalStageList.RemoveAt(index);
                            LocalStageList.Insert(index + 1, stritem);
                            smallPopulateStage();
                            costStageLv.Items[index + 1].Selected = true;
                            index = index + 1;
                            break;
                        }
                    }
                }
            }
        }
        private void CalculateClist()
        {  
            LocalCostList.Clear();
            foreach (CSstage Stageditem in LocalStageList)
            {

                double packsAmnt = Stageditem.packsAmnt;
                double Ammount = Stageditem.Ammount;
                string symbol = Stageditem.symbol;
                string inputname = Stageditem.inputname;
                double cost = Stageditem.cost;
                double Packvol = Stageditem.Packvol;
                double price = Stageditem.price;

                bool wasinList = false;
                foreach (Costing Costitem in LocalCostList)
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
                    LocalCostList.Add(nc);
                }
            }
        }

        private void ExportPobtn_Click(object sender, EventArgs e)
        {            
            if (LocalStageList.Count > 0)
            {
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                object start_of_doc = "\\startofdoc";

                List<CSstage> cstgList = new List<CSstage>();
                cstgList.Clear();
                List<string> Lslist = new List<string>();
                List<string> LsAreaList = new List<string>();
                List<string> LsStagePriceList = new List<string>();
                foreach (CSstage stgitem in LocalStageList)
                {                    
                    cstgList.Add(stgitem);                    
                    string currentstring = stgitem.Staging;
                    if (!Lslist.Contains(currentstring))
                    {
                        Lslist.Add(stgitem.Staging); 
                        if (stgitem.Area > 0 && stgitem.Lm <= 0)
                        {
                            LsAreaList.Add(stgitem.Area.ToString() + " m2");
                        }
                        else if (stgitem.Lm > 0 && stgitem.Area <= 0)
                        {
                            LsAreaList.Add(stgitem.Lm.ToString() + " Lm");

                        }
                        else if (stgitem.Area > 0 && stgitem.Lm > 0)
                        {
                            if (stgitem.Systems == "Subfill")
                            {
                                LsAreaList.Add(stgitem.Area.ToString() + " m2, " + stgitem.Lm.ToString() + " mm");
                            }
                            else
                            {
                                LsAreaList.Add(stgitem.Area.ToString() + " m2, " + stgitem.Lm.ToString() + " Lm");
                            }
                        }
                    }
                }

                string[,] CstListArray;
                CstListArray = new string[LocalStageList.Count + 2 + Lslist.Count*2, 8];
                int i = 1;
                int j = 0;
                double costpr = 0;
                foreach (string stritem in Lslist)
                {                    
                    //SummaryLB.Items.Add("---------------- " + stritem + " ----------------");
                    i++;
                    CstListArray[i, 1] = stritem;
                    CstListArray[i, 2] = LsAreaList[j];
                    CstListArray[i, 3] = "";
                    CstListArray[i, 4] = "";
                    CstListArray[i, 5] = "";
                    CstListArray[i, 6] = "";
                    CstListArray[i, 7] = "";
                    //CstListArray[i, 8] = "";
                    //CstListArray[i, 7] = "";                    
                    CSstage newcst = new CSstage();
                    newcst.Staging = stritem;
                    newcst.Systems = "";
                    newcst.inputname = "";
                    cstgList.Add(newcst);
                    foreach (CSstage Cstgitem in LocalStageList)
                    {                        
                        if (stritem == Cstgitem.Staging)
                        {
                            costpr += Cstgitem.price;
                            //LsStagePriceList.Add(costpr.ToString());
                            i++;
                            cstgList.Add(Cstgitem);
                            CstListArray[i, 1] = "";
                            CstListArray[i, 2] = Cstgitem.inputname;
                            CstListArray[i, 3] = Math.Round(Cstgitem.Ammount, 2).ToString() + " " + Cstgitem.symbol;
                            CstListArray[i, 4] = Cstgitem.Packvol.ToString() + " " + Cstgitem.symbol;
                            CstListArray[i, 5] = Math.Ceiling(Cstgitem.packsAmnt).ToString();
                            CstListArray[i, 6] = "$ " + Math.Round(Cstgitem.price, 2).ToString();
                            CstListArray[i, 7] = "";
                            //CstListArray[i, 8] = "";
                            //CstListArray[i, 3] = "$" + Math.Round(Coitem.cost, 2).ToString();
                            //CstListArray[i, 4] = "$" + Math.Round(Coitem.price, 2).ToString();

                        }
                    }
                    i++;
                    CstListArray[i, 1] = "";
                    CstListArray[i, 2] = "";
                    CstListArray[i, 3] = "";
                    CstListArray[i, 4] = "";
                    CstListArray[i, 5] = "";
                    CstListArray[i, 6] = "";
                    CstListArray[i, 7] = "$ " + Math.Round(costpr,2).ToString();// LsStagePriceList[j];

                    CSstage newcst2 = new CSstage();
                    newcst2.Staging = "";
                    newcst2.Systems = "";
                    newcst2.inputname = "";
                    newcst2.price = costpr;
                    cstgList.Add(newcst2);
                    costpr = 0;
                    j++;
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
                string path = wanted_path + "..\\..\\Data\\Qepoxy1.jpg";
                range.Application.ActiveDocument.InlineShapes.AddPicture(@path, ref oLinkToFile, ref oSaveWithDocument, ref oMissed);

                Microsoft.Office.Interop.Word.Shape s = oDoc.Shapes.AddShape(1, 320, 100, 250, 35);
                s.Fill.ForeColor.RGB = Color.FromArgb(186, 174, 255).ToArgb();
                s.TextFrame.TextRange.Text = "Statement";
                s.TextFrame.TextRange.Font.Size = 24;
                s.TextFrame.TextRange.Font.Name = "Arial";
                s.TextFrame.TextRange.Font.Color = 0;

                int spacing = 12;
                int Titlefontsize = 11;
                //Word.Paragraph oPara3z;
                //oPara3z = oDoc.Content.Paragraphs.Add(ref oMissing);
                //oPara3z.Format.SpaceAfter = spacing;
                //oPara3z.Range.InsertParagraphAfter();

                Word.Paragraph oPara1a;
                oPara1a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara1a.Range.Text = "Product Sheet";
                oPara1a.Range.Font.Bold = 1;
                oPara1a.Range.Font.Size = Titlefontsize;
                oPara1a.Range.Font.Name = "calibri";
                oPara1a.Format.SpaceAfter = spacing;

                oPara1a.Range.InsertParagraphAfter();
                Word.Paragraph oPara2a;
                oPara2a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara2a.Range.Text = "Quote #: " + JsizeTb.Text;
                oPara2a.Range.Font.Bold = 1;
                oPara2a.Range.Font.Size = Titlefontsize;
                oPara2a.Range.Font.Name = "calibri";
                oPara2a.Format.SpaceAfter = spacing;
                oPara2a.Range.InsertParagraphAfter();

                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3a.Range.Text = "Job Name: " + JnameTb.Text;
                oPara3a.Range.Font.Bold = 1;
                oPara3a.Range.Font.Size = Titlefontsize;
                oPara3a.Range.Font.Name = "calibri";
                oPara3a.Format.SpaceAfter = spacing;
                oPara3a.Range.InsertParagraphAfter();

                Word.Paragraph oPara3c;
                oPara3c = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3c.Range.Text = "Job #: " + Jobnumtb.Text;
                oPara3c.Range.Font.Bold = 1;
                oPara3c.Range.Font.Size = Titlefontsize;
                oPara3c.Range.Font.Name = "calibri";
                oPara3c.Format.SpaceAfter = spacing;
                oPara3c.Range.InsertParagraphAfter();

                Word.Paragraph oPara3b;
                oPara3b = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3b.Range.Text = "Job Size: " + JsizeTb.Text;
                oPara3b.Range.Font.Bold = 1;
                oPara3b.Range.Font.Size = Titlefontsize;
                oPara3b.Range.Font.Name = "calibri";
                oPara3b.Format.SpaceAfter = spacing;
                oPara3b.Range.InsertParagraphAfter();

                Word.Paragraph oPara2;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
                //int tablerows = cstgList.Count- Lslist.Count;
                int tablerows = i;
                Word.Table firstTable = oDoc.Tables.Add(oPara2.Range, tablerows, 7, ref oMissing, ref oMissing);
                int Basefontsize = 9;
                firstTable.Borders.Enable = 0;
                //float width = 20;
                firstTable.Columns[2].SetWidth(140, Word.WdRulerStyle.wdAdjustSameWidth);
                firstTable.Columns[3].SetWidth(70, Word.WdRulerStyle.wdAdjustSameWidth);
                firstTable.Columns[7].SetWidth(70, Word.WdRulerStyle.wdAdjustSameWidth);
                //firstTable.Rows.SetHeight(8, Word.WdRowHeightRule.wdRowHeightAuto);
                //firstTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[2].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[3].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[4].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[5].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[6].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                firstTable.Columns[7].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                // firstTable.Columns[8].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                foreach (Word.Row row in firstTable.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        if (cell.ColumnIndex == 1)
                        {

                        }
                        //Header row  
                        if (cell.RowIndex == 1 && cell.ColumnIndex == 1)
                        {
                            cell.Range.Text = "Stage ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
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
                            cell.Range.Text = "Description: ";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 3)
                        {

                            cell.Range.Text = "Qty:";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
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

                            cell.Range.Text = "Size:";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 5)
                        {
                            cell.Range.Text = "#:";

                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 6)
                        {
                            cell.Range.Text = "$Price val";
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
                            cell.Range.Paragraphs.SpaceBefore = 0;
                            cell.Range.Paragraphs.SpaceAfter = 0;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        else if (cell.RowIndex == 1 && cell.ColumnIndex == 7)
                        {
                            cell.Range.Text = "Stage-Cost:";

                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize + 2;
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
                            //if (cell.ColumnIndex == 6)
                            //{
                            //    //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                            //    cell.Range.Font.Bold = 1;
                            //    //other format properties goes here  
                            //    firstTable.Cell(cell.RowIndex, cell.ColumnIndex).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            //    cell.Range.Font.Name = "calibri";
                            //    cell.Range.Font.Size = 12;
                            //    cell.Range.Paragraphs.SpaceBefore = 0;
                            //    cell.Range.Paragraphs.SpaceAfter = 0;
                            //    cell.Range.Text = CstListArray[cell.RowIndex, cell.ColumnIndex];

                            //}
                            //else
                            //{
                                cell.Range.Font.Bold = 0;
                                //firstTable.Cell(cell.RowIndex, cell.ColumnIndex).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize;
                                //cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                cell.Range.Text = CstListArray[cell.RowIndex, cell.ColumnIndex];
                                cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            //}
                        }
                        else
                        {
                            //cell.Range.Text = "$" + (Subtotal * 1.1).ToString();
                            //cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "calibri";
                            cell.Range.Font.Size = Basefontsize;
                            //cell.Range.Paragraphs.SpaceBefore = 0;
                            //cell.Range.Paragraphs.SpaceAfter = 0;                            
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                            //Center alignment for the Header cells  
                            //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            //cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                }
                if(Additionscb.Checked)
                {
                    string[,] AddsListArray;
                    AddsListArray = new string[newBe.lists.AddsList.Count + 2, 5];
                    int n = 1;
                    foreach (Adds Additem in newBe.lists.AddsList)
                    {
                        n++;
                        AddsListArray[n, 0] = "";
                        AddsListArray[n, 1] = Additem.ItemName;
                        AddsListArray[n, 2] = Additem.Desc;
                        AddsListArray[n, 3] = Additem.Ammount;
                        AddsListArray[n, 4] = "$" + Additem.Cost;
                    }

                    Word.Paragraph oPara33a;
                    oPara33a = oDoc.Content.Paragraphs.Add(ref oMissing);
                    //oPara3a.Range.Text = "Total: $ " + Math.Round(double.Parse(EstimTb.Text), 2).ToString();
                    //oPara33a.Range.Font.Bold = 1;
                    //oPara33a.Range.Font.Size = Titlefontsize + 5;
                    //oPara33a.Range.Font.Name = "calibri";
                    //oPara33a.Format.SpaceAfter = spacing;
                    //oPara33a.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara33a.Range.InsertParagraphAfter();

                    Word.Paragraph oPara3;
                    oPara3 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    //int tablerows = cstgList.Count- Lslist.Count;
                    int addstablerows = n;
                    Word.Table secondTable;
                    secondTable = oDoc.Tables.Add(oPara3.Range, addstablerows, 4, ref oMissing, ref oMissing);
                    secondTable.Borders.Enable = 1;
                    //float width = 20;
                    secondTable.Columns[1].SetWidth(140, Word.WdRulerStyle.wdAdjustNone);
                    secondTable.Columns[2].SetWidth(180, Word.WdRulerStyle.wdAdjustNone);
                    secondTable.Columns[4].SetWidth(40, Word.WdRulerStyle.wdAdjustNone);
                    //secondTable.Rows.SetHeight(8, Word.WdRowHeightRule.wdRowHeightAuto);
                    secondTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    secondTable.Columns[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    secondTable.Columns[2].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    secondTable.Columns[3].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    secondTable.Columns[4].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    foreach (Word.Row row in secondTable.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            if (cell.ColumnIndex == 1)
                            {

                            }
                            //Header row  
                            if (cell.RowIndex == 1 && cell.ColumnIndex == 1)
                            {
                                cell.Range.Text = "Item Name:";
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize + 2;
                                cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                                //Center alignment for the Header cells  
                                //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            if (cell.RowIndex == 1 && cell.ColumnIndex == 2)
                            {
                                cell.Range.Text = "Description:";
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize + 2;
                                cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                                //Center alignment for the Header cells  
                                //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            if (cell.RowIndex == 1 && cell.ColumnIndex == 3)
                            {
                                cell.Range.Text = "Amount:";
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize + 2;
                                cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                                //Center alignment for the Header cells  
                                //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            if (cell.RowIndex == 1 && cell.ColumnIndex == 4)
                            {
                                cell.Range.Text = "Cost:";
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize + 2;
                                cell.Range.Paragraphs.SpaceBefore = 0;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                                //Center alignment for the Header cells  
                                //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            if (cell.RowIndex > 1 && cell.RowIndex <= tablerows)
                            {
                                cell.Range.Font.Bold = 0;
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize;
                                cell.Range.Paragraphs.SpaceAfter = 0;
                                cell.Range.Text = AddsListArray[cell.RowIndex, cell.ColumnIndex];
                                cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                //}
                            }
                            else
                            {
                                //cell.Range.Text = "$" + (Subtotal * 1.1).ToString();
                                //cell.Range.Font.Bold = 1;
                                //other format properties goes here  
                                cell.Range.Font.Name = "calibri";
                                cell.Range.Font.Size = Basefontsize;
                                //cell.Range.Paragraphs.SpaceBefore = 0;
                                //cell.Range.Paragraphs.SpaceAfter = 0;                            
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                                //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                                //Center alignment for the Header cells  
                                //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                //cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }
                    }
                    Word.Paragraph oPara34a;
                    oPara34a = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara34a.Range.InsertParagraphAfter();
                    double addoutput = 0;
                    foreach (Adds Additem in newBe.lists.AddsList)
                    {
                        addoutput += double.Parse(Additem.Cost);
                    }

                    Word.Paragraph oPara34b;
                    oPara34b = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara34b.Range.Text = "Additions: $ " + Math.Round(addoutput, 2).ToString();
                    oPara34b.Range.Font.Bold = 1;
                    oPara34b.Range.Font.Size = Titlefontsize + 5;
                    oPara34b.Range.Font.Name = "calibri";
                    oPara34b.Format.SpaceAfter = spacing;
                    oPara34b.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara34b.Range.InsertParagraphAfter();
                }
                

                Word.Paragraph oPara4a;
                oPara4a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara4a.Range.Text = "Total: $ " + Math.Round(double.Parse(EstimTb.Text),2).ToString();
                oPara4a.Range.Font.Bold = 1;
                oPara4a.Range.Font.Size = Titlefontsize+5;
                oPara4a.Range.Font.Name = "calibri";
                oPara4a.Format.SpaceAfter = spacing;
                oPara4a.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara4a.Range.InsertParagraphAfter();

                foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
                {
                    Word.Range header = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    Word.Paragraph h1 = header.Paragraphs.Add(header);
                    header.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    header.Text = "Version: 2.5";
                    header.Font.Size = 8;
                    header.Font.Bold = 0;
                    // header.InsertParagraphAfter();

                    Word.Range footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    Word.Paragraph p1 = footer.Paragraphs.Add(footer);
                    footer.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //footer.Text = "Version: 2.5";
                    footer.Font.Size = 8;
                    footer.Font.Bold = 0;
                    Word.Paragraph p3 = footer.Paragraphs.Add(footer);
                    p3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    p3.Range.Fields.Add(footer , Word.WdFieldType.wdFieldPage);
                }


            }
        }
    }
}
