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
using System.Reflection;
using System.Drawing.Imaging;

namespace Qepoxy
{
    public partial class PurchaseOrder : Form
    {
        List<Costing> cstList = new List<Costing>();
        string[,] CstListArray;
        Backend newBe;
        double Subtotal = 0;
        public PurchaseOrder(Backend newBackend)
        {
            newBe = newBackend;
            InitializeComponent();

            this.PoCostingslv.Scrollable = true;
            this.PoCostingslv.View = View.Details;
            this.PoCostingslv.GridLines = true;
            this.PoCostingslv.FullRowSelect = true;
            this.PoCostingslv.Columns.Add("Product", 160);
            this.PoCostingslv.Columns.Add("Qty", 50);        
            this.PoCostingslv.Columns.Add("Unit Price", 80);
            this.PoCostingslv.Columns.Add("$ Amount", 80);

            this.Productlv.Scrollable = true;
            this.Productlv.View = View.Details;
            this.Productlv.GridLines = true;
            this.Productlv.FullRowSelect = true;
            this.Productlv.Columns.Add("Product", 220);
            this.Productlv.Columns.Add("Price", 50);
            this.Productlv.Columns.Add("Weight", 40);
            this.Productlv.Columns.Add("Volume", 40);
            this.Productlv.Columns.Add("Coverage", 40);
            this.Productlv.Columns.Add("Unit", 40);

            this.Unitcb.Items.AddRange(new object[] { "L" });
            this.Unitcb.Items.AddRange(new object[] { "Kg" });
            this.Unitcb.Items.AddRange(new object[] { "m" });
            this.Unitcb.Items.AddRange(new object[] { "Pack" });
            this.Unitcb.Items.AddRange(new object[] { "Bag" });

            foreach (Costing Coitem in newBe.lists.POCostList)
            {
                if (Coitem.inputname == "Hours cost")
                {
                    newBe.lists.POCostList.Remove(Coitem);
                    break;
                }
            }
            CreateList();
        }
        private void CreateList()
        {
            cstList.Clear();
            CstListArray = new string[newBe.lists.POCostList.Count + 2, 5];
            int i = 1;
            foreach (Costing Coitem in newBe.lists.POCostList)
            {
                i++;
                cstList.Add(Coitem);
                CstListArray[i, 1] = Coitem.inputname;
                CstListArray[i, 2] = Math.Round(Coitem.Ammount, 2).ToString() + " " + Coitem.symbol;                
                CstListArray[i, 3] = "$" + Math.Round(Coitem.cost, 2).ToString();
                CstListArray[i, 4] = "$" + Math.Round(Coitem.price, 2).ToString();
            }
            Populate();
        }
        private void Populate()
        {
            Subtotal = 0;
            this.PoCostingslv.Items.Clear();
            foreach (Costing Coitem in newBe.lists.POCostList)
            {
                    Subtotal += Math.Round(Coitem.price, 2);
                    string[] newArr = new string[12];
                    newArr[0] = Coitem.inputname;
                    newArr[1] =  Math.Round(Coitem.Ammount,2).ToString() + " " + Coitem.symbol;                   
                    newArr[2] = "$" + Math.Round(Coitem.cost,2).ToString();
                    newArr[3] = "$" + Math.Round(Coitem.price, 2).ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    this.PoCostingslv.Items.Add(itm);
            }

            this.Productlv.Items.Clear();
            foreach(Priceitem Pitem in newBe.lists.PriceList)
            {
                string[] newArr = new string[12];
                newArr[0] = Pitem.Name;
                newArr[1] = Pitem.Price.ToString();
                newArr[2] = Pitem.Weight.ToString();
                newArr[3] = Pitem.Volume.ToString();
                newArr[4] = Pitem.Coverage.ToString();
                newArr[5] = Pitem.Unit;
                ListViewItem itm = new ListViewItem(newArr);
                this.Productlv.Items.Add(itm);
            }
        }

        private void POgenbtn_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            object start_of_doc = "\\startofdoc";

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
            s.Fill.ForeColor.RGB = Color.FromArgb(81, 208, 146).ToArgb();
            s.TextFrame.TextRange.Text = "Purchase Order";
            s.TextFrame.TextRange.Font.Size = 24;
            s.TextFrame.TextRange.Font.Name = "Arial";
            s.TextFrame.TextRange.Font.Color = 0;

            //Footer:

            foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
            {
                Word.Range footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                Word.Paragraph p1 = footer.Paragraphs.Add(footer);
                footer.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footer.Text = "Jacob Young - Qepoxy Pty Ltd";
                footer.Font.Size = 10;
                footer.Font.Bold = 1;
                footer.InsertParagraphAfter();

                Word.Paragraph p2 = footer.Paragraphs.Add(p1.Range);
                p2.Range.Text = "ABN: 33 628 159 428 | QBCC: 15140664";
                p2.Range.Font.Size = 10;
                p2.Range.Font.Bold = 0;
                p2.Range.InsertParagraphAfter();

                Word.Paragraph p3 = footer.Paragraphs.Add(p2.Range);
                p3.Range.Text = "Jacob@qepoxy.com.au | Mobile: 0447 071 047 | Ph: 1800 737 699";
                p3.Range.InsertParagraphAfter();

                Word.Paragraph p4 = footer.Paragraphs.Add(p3.Range);
                p4.Range.Text = "www.qepoxy.com.au";
                p4.Range.InsertParagraphAfter();

            }

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
            oPara1a.Range.Text = "Attention: " + Attentb.Text + "\t\t\t\t\t\t Date: "+ Datetb.Text;
            oPara1a.Range.Font.Bold = 1;
            oPara1a.Range.Font.Size = 12;
            oPara1a.Range.Font.Name = "calibri";
            oPara1a.Format.SpaceAfter = spacing;
            oPara1a.Range.InsertParagraphAfter();

            Word.Paragraph oPara1b;
            oPara1b = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1b.Range.Text = "Company: "+ Companytb.Text + "\t\t\t\t\t PO Number: " + PoNumtb.Text;
            oPara1b.Range.Font.Bold = 1;
            oPara1b.Range.Font.Size = 12;
            oPara1b.Range.Font.Name = "calibri";
            oPara1b.Format.SpaceAfter = spacing;
            oPara1b.Range.InsertParagraphAfter();

            Word.Paragraph oPara1c;
            oPara1c = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1c.Range.Text = "Email: " + Emailtb.Text + "\t\t\t\t\t\t\t Delivery: " + Deliverytb.Text;
            oPara1c.Range.Font.Bold = 1;
            oPara1c.Range.Font.Size = 12;
            oPara1c.Range.Font.Name = "calibri";
            oPara1c.Format.SpaceAfter = spacing + 12;    //20 pt spacing after paragraph.
            oPara1c.Range.InsertParagraphAfter();

            Word.Paragraph oPara2;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
            int tablerows = cstList.Count+1;
            Word.Table firstTable = oDoc.Tables.Add(oPara2.Range, tablerows+4, 4, ref oMissing, ref oMissing);

            firstTable.Borders.Enable = 1;
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
                    else if(cell.RowIndex <= tablerows)
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
                    else if(cell.RowIndex == tablerows+2 && cell.ColumnIndex == 3)
                    {
                        cell.Range.Text = "Subtotal ";
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (cell.RowIndex == tablerows + 2 && cell.ColumnIndex == 4)
                    {
                        cell.Range.Text = "$"+Subtotal.ToString();
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (cell.RowIndex == tablerows + 3 && cell.ColumnIndex == 3)
                    {
                        cell.Range.Text = "GST ";
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (cell.RowIndex == tablerows + 3 && cell.ColumnIndex == 4)
                    {
                        cell.Range.Text = "$" + (Subtotal*0.1).ToString();
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (cell.RowIndex == tablerows + 4 && cell.ColumnIndex == 3)
                    {
                        cell.Range.Text = "Total ";
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else if (cell.RowIndex == tablerows + 4 && cell.ColumnIndex == 4)
                    {
                        cell.Range.Text = "$" + (Subtotal * 1.1).ToString();
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here  
                        cell.Range.Font.Name = "calibri";
                        cell.Range.Font.Size = 10;
                        cell.Range.Paragraphs.SpaceBefore = 0;
                        cell.Range.Paragraphs.SpaceAfter = 0;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                        //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        //Center alignment for the Header cells  
                        //cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string today = DateTime.Now.ToString("dd/MM/yyyy");
            this.Datetb.Text = today;
        }

        private void Addbtn_Click(object sender, EventArgs e)
        {
            double Qnty = 0;           
            bool Qntypassed = double.TryParse(Quantitytb.Text, out Qnty);            
            if (Qntypassed && Unitcb.Text != "")
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
                            double ratio = 0;
                            Costing newCstitm = new Costing();
                            newCstitm.inputname = PItem.Name;
                            if (PItem.Volume == 0)
                            {
                                ratio = Qnty / PItem.Weight;
                                newCstitm.Packvol = PItem.Weight;
                                newCstitm.symbol = Unitcb.Text;
                                newCstitm.Ammount = PItem.Weight * ratio;
                            }
                            else if (PItem.Weight == 0)
                            {
                                ratio = Qnty / PItem.Volume;
                                newCstitm.Packvol = PItem.Volume;
                                newCstitm.symbol = Unitcb.Text;
                                newCstitm.Ammount = ratio * PItem.Volume;
                            }
                            else
                            {
                                ratio = Qnty / PItem.Volume;
                                newCstitm.Packvol = PItem.Volume;
                                newCstitm.symbol = Unitcb.Text;
                                newCstitm.Ammount = ratio * PItem.Volume;
                            }
                            newCstitm.cost = PItem.Price;
                            newCstitm.price = ratio * PItem.Price;
                            newBe.lists.POCostList.Add(newCstitm);
                        }
                    }
                    CreateList();
                }
            }
        }

        private void Rembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (PoCostingslv.SelectedIndices.Count > 0)
            {
                Selected = PoCostingslv.SelectedIndices[0];
                string item = PoCostingslv.Items[Selected].Text;
                foreach (Costing CItem in newBe.lists.POCostList)
                {
                    if (CItem.inputname == item)
                    {
                        newBe.lists.POCostList.Remove(CItem);
                        break;
                    }
                }
                CreateList();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Jobnum f4 = new Jobnum();
            f4.Owner = this;
            f4.Show();
        }

        private void Editbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            int index = 0;
            Costing CItemEdit = new Costing();
            if (PoCostingslv.SelectedIndices.Count > 0)
            {
                Selected = PoCostingslv.SelectedIndices[0];
                string item = PoCostingslv.Items[Selected].Text;
                foreach (Costing CItem in newBe.lists.POCostList)
                {
                    if (CItem.inputname == item)
                    {
                        CItemEdit = CItem;
                        index = newBe.lists.POCostList.IndexOf(CItem);
                        break;
                    }
                }                
            }

            POItemEdit Poi = new POItemEdit(index, CItemEdit.inputname, CItemEdit.Ammount,CItemEdit.cost, CItemEdit.Packvol,CItemEdit.symbol);
            Poi.Owner = this;
            Poi.Show();
        }

        public void SaveItemTolist(int index, string Desc, double Qty, double Cost, double Price,double Packamnt, string Unit)
        {
            Costing CItem = new Costing();
            CItem.inputname = Desc;
            CItem.cost = Cost;
            CItem.price = Price;
            CItem.Ammount = Qty;
            CItem.Packvol = Packamnt;
            CItem.symbol = Unit;
            newBe.lists.POCostList[index] = CItem;
            CreateList();
        }
    }
}
