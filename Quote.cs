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
    public partial class Quote : Form
    {
        Backend newBe;
        List<QuoteComment> GenCommentList = new List<QuoteComment>();

        //List<Company> CompList = new List<Company>();
        string QuoteName;
        List<string> NewCommentList = new List<string>();

        public Quote(Backend Backend, double InEstcost,string InQuoteName)
        {
            InitializeComponent();
            //LoadDbFromDisk();
            //LoadCompanyData();
            //LoadGCommentsDb();

            newBe = Backend;
            Jpricetb.Text = InEstcost.ToString();
            this.NewCommentLv.Scrollable = true;
            this.NewCommentLv.View = View.Details;
            this.NewCommentLv.GridLines = true;
            this.NewCommentLv.FullRowSelect = true;
            this.NewCommentLv.Columns.Add("Comments", 650);

            this.GCdbLv.Scrollable = true;
            this.GCdbLv.View = View.Details;
            this.GCdbLv.GridLines = true;
            this.GCdbLv.FullRowSelect = true;
            this.GCdbLv.Columns.Add("Comments", 450);

            this.GcomLv.Scrollable = true;
            this.GcomLv.View = View.Details;
            this.GcomLv.GridLines = true;
            this.GcomLv.FullRowSelect = true;
            this.GcomLv.Columns.Add("Comments", 450);

            this.CompanyLv.Scrollable = true;
            this.CompanyLv.View = View.Details;
            this.CompanyLv.GridLines = true;
            this.CompanyLv.FullRowSelect = true;
            this.CompanyLv.Columns.Add("Company", 270);

            this.CommentDblv.Scrollable = true;
            this.CommentDblv.View = View.Details;
            this.CommentDblv.GridLines = true;
            this.CommentDblv.FullRowSelect = true;
            this.CommentDblv.Columns.Add("Headings", 480);
            this.CommentDblv.Columns.Add("#Comments", 40);

            this.Add2Quote1lv.Scrollable = true;
            this.Add2Quote1lv.View = View.Details;
            this.Add2Quote1lv.GridLines = true;
            this.Add2Quote1lv.FullRowSelect = true;
            this.Add2Quote1lv.Columns.Add("Headings", 350);
            this.Add2Quote1lv.Columns.Add("#Comments", 50);

            this.Add2Quote2lv.Scrollable = true;
            this.Add2Quote2lv.View = View.Details;
            this.Add2Quote2lv.GridLines = true;
            this.Add2Quote2lv.FullRowSelect = true;
            this.Add2Quote2lv.Columns.Add("Headings", 350);
            this.Add2Quote2lv.Columns.Add("#Comments", 50);

            this.Add2Quote3lv.Scrollable = true;
            this.Add2Quote3lv.View = View.Details;
            this.Add2Quote3lv.GridLines = true;
            this.Add2Quote3lv.FullRowSelect = true;
            this.Add2Quote3lv.Columns.Add("Headings", 350);
            this.Add2Quote3lv.Columns.Add("#Comments", 50);

            this.QuotelistLv.Scrollable = true;
            this.QuotelistLv.View = View.Details;
            this.QuotelistLv.GridLines = true;
            this.QuotelistLv.FullRowSelect = true;
            this.QuotelistLv.Columns.Add("QuoteName", 280);
            //this.QuotelistLv.Columns.Add("Save Data", 80);


            this.Aheadbox.Text = "Us";
            this.Aheadbox.Items.AddRange(new object[] { "Us" });
            this.Aheadbox.Items.AddRange(new object[] { "You" });

            this.Hstatecb.Text = "";
            this.Hstatecb.Items.AddRange(new object[] { "ACT" });            
            this.Hstatecb.Items.AddRange(new object[] { "NSW" });
            this.Hstatecb.Items.AddRange(new object[] { "NT" });
            this.Hstatecb.Items.AddRange(new object[] { "QLD" });
            this.Hstatecb.Items.AddRange(new object[] { "SA" });
            this.Hstatecb.Items.AddRange(new object[] { "TAS" });
            this.Hstatecb.Items.AddRange(new object[] { "VIC" });
            this.Hstatecb.Items.AddRange(new object[] { "WA" });

            this.CompStatecb.Text = "";
            this.CompStatecb.Items.AddRange(new object[] { "ACT" });            
            this.CompStatecb.Items.AddRange(new object[] { "NSW" });
            this.CompStatecb.Items.AddRange(new object[] { "NT" });
            this.CompStatecb.Items.AddRange(new object[] { "QLD" });
            this.CompStatecb.Items.AddRange(new object[] { "SA" });
            this.CompStatecb.Items.AddRange(new object[] { "TAS" });
            this.CompStatecb.Items.AddRange(new object[] { "VIC" });
            this.CompStatecb.Items.AddRange(new object[] { "WA" });


            Warrantytb.Enabled = false;
            Completetb.Enabled = false;
            QuoteName = InQuoteName;
            Job3Headtb.Enabled = false;
            Add2Quote3lv.Enabled = false;
            Add2Quote3btn.Enabled = false;
            RemCmntFrmQuote3btn.Enabled = false;
            Job2m2tb.Enabled = false;
            Job2lmtb.Enabled = false;
            Job2Comptb.Enabled = false;
            H1Warrant2Tb.Enabled = false;
            J2Pricetb.Enabled = false;
            J2slipratetb.Enabled = false;

            Job2Headtb.Enabled = false;
            Add2Quote2lv.Enabled = false;
            Add2Quote2btn.Enabled = false;
            RemCmntFrmQuote2btn.Enabled = false;
            Job3m2tb.Enabled = false;
            Job3lmtb.Enabled = false;
            Job3Comptb.Enabled = false;
            H1Warrant3Tb.Enabled = false;
            J3Pricetb.Enabled = false;
            
            J3slipratetb.Enabled = false;
            StatusLabel1.Text = "";
            this.FormClosing += Form_FormClosing;
            Populate();
            populateCdb();
            PopulateGenCommentDb();
            populateCompanyList();

            //newBe.lists.RebuildQuoteDBLists();

        }
        private void Form_FormClosing(Object sender, FormClosingEventArgs e)
        {
            //SaveDbToDisk();
            //SaveCompanyData();
            //SaveGCommentsDb();
        }
        private void loadQuotes()
        {
            int Selected = 0;
            if (QuotelistLv.SelectedIndices.Count > 0)
            {
                Selected = QuotelistLv.SelectedIndices[0];
                string item = QuotelistLv.Items[Selected].Text;
                QuoteNametb.Text = item;                
            }
        }
        
        public void GenerateGeneralComment()
        {
            QuoteComment Qcomment1 = new QuoteComment();
            GenCommentList.Clear();
            Qcomment1.QcHeading = "General:";
            if(Rubbcb.Checked)
                Qcomment1.QcTextList.Add("Rubbish removal (by us) into site bins (provided by you)");
            if(PH3cb.Checked)
                Qcomment1.QcTextList.Add("3 phase power will be supplied by YOU");
            if (DryTimecb.Checked)
                Qcomment1.QcTextList.Add("Min 24 hours dry-time is required after final sealer before access (foot traffic)");
            if (DryTimecb.Checked)
                Qcomment1.QcTextList.Add("Min 48 hours dry-time before any heavy vehicle traffic");
            if (QuoteStagecb.Checked)
            {
                double num1 = 0;
                if(QuoteStagestb.Text != "")
                    num1 = double.Parse(QuoteStagestb.Text);

                Qcomment1.QcTextList.Add("This quotation has allowed for works to be completed over one (" + num1.ToString() + ") stage");
            }
            if (Jointcb.Checked)
            {
                Qcomment1.QcTextList.Add("Jointing compound colour is "+ JCCtb.Text);
            }
            if (FloorColourcb.Checked)
            {
                Qcomment1.QcTextList.Add("Colour to be " + Coltb.Text);
            }
            if (SlipRcb.Checked)
                Qcomment1.QcTextList.Add("Slip resistance – " + SRtb.Text);
            if (Aheadcb.Checked)
                Qcomment1.QcTextList.Add("Arrow-head (perimeter for coving) to be installed by " + Aheadbox.Text);
            if (ContJcb.Checked)
                Qcomment1.QcTextList.Add("Allowance for control joints NOT been made - $" + ContJtb.Text + " per Lm if required");


            foreach(string Gcomment in newBe.lists.GCommentList)
            {
                Qcomment1.QcTextList.Add(Gcomment);
            }

            //if (Fallsdrainscb.Checked)
            //    Qcomment1.QcTextList.Add("Allowance for falls to drains has been made");
            //else
            //    Qcomment1.QcTextList.Add("Allowance for falls to drains has NOT been made");
            //if (Hrsworkcb.Checked)
            //    Qcomment1.QcTextList.Add("Allowance for out of hours works (Night Shift/Weekends) has been made");
            //else
            //    Qcomment1.QcTextList.Add("Allowance for out of hours works (Night Shift/Weekends) has NOT been made");
            //if (AwayWcb.Checked)
            //    Qcomment1.QcTextList.Add("Allowance for away works/overnight allowance has been made");
            //else
            //    Qcomment1.QcTextList.Add("Allowance for away works/overnight allowance NOT been made");
            //if (UVstabcb.Checked)
            //    Qcomment1.QcTextList.Add("Allowance for UV stable topcoat has been made");
            //else
            //    Qcomment1.QcTextList.Add("Allowance for UV stable topcoat has NOT been made");
            //if (AccelFloorcb.Checked)
            //    Qcomment1.QcTextList.Add("Allowance for accelerated flooring has been made");
            //else
            //    Qcomment1.QcTextList.Add("Allowance for accelerated flooring has NOT been made");
            GenCommentList.Add(Qcomment1);
        }

        public void populateCompanyList()
        { 
            CompanyLv.Items.Clear();       
            newBe.lists.RebuildCompanyList();
            foreach (Company cI in newBe.lists.Complist)
            {                
                CompanyLv.Items.Add(cI.CompName);
            }
        }

        public void Populate()
        {

            QuotelistLv.Items.Clear();            
            foreach (Quote_Block Qfsitem in newBe.lists.Q_bList)
            {
                if (Qfsitem.QName.ToUpper().Contains(SearchTb.Text.ToUpper()))
                {
                    string[] newArr = new string[40];
                    newArr[0] = Qfsitem.QName;
                    ListViewItem itm = new ListViewItem(newArr);
                    QuotelistLv.Items.Add(itm);
                }
                else if (SearchTb.Text == "")
                {
                    string[] newArr = new string[40];
                    newArr[0] = Qfsitem.QName;
                    ListViewItem itm = new ListViewItem(newArr);
                    QuotelistLv.Items.Add(itm);
                }
            }

            NewCommentLv.Items.Clear();
            foreach (string Cmnt in NewCommentList)
            {
                NewCommentLv.Items.Add(Cmnt);
            }           

            Add2Quote1lv.Items.Clear();
            foreach (QuoteComment CmntQ in newBe.lists.Add2QuoteList)
            {
                string[] newArr = new string[5];
                newArr[0] = CmntQ.QcHeading;
                newArr[1] = CmntQ.QcTextList.Count.ToString();                
                if (CmntQ.QcTextList.Count > 0)
                { 
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Add2Quote1lv.Items.Add(itm);
                }
            }

            Add2Quote2lv.Items.Clear();
            foreach (QuoteComment CmntQ in newBe.lists.Add2QuoteList2)
            {
                string[] newArr = new string[5];
                newArr[0] = CmntQ.QcHeading;
                newArr[1] = CmntQ.QcTextList.Count.ToString();
                if (CmntQ.QcTextList.Count > 0)
                {
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Add2Quote2lv.Items.Add(itm);
                }
            }

            Add2Quote3lv.Items.Clear();
            foreach (QuoteComment CmntQ in newBe.lists.Add2QuoteList3)
            {
                string[] newArr = new string[5];
                newArr[0] = CmntQ.QcHeading;
                newArr[1] = CmntQ.QcTextList.Count.ToString();
                if (CmntQ.QcTextList.Count > 0)
                {
                    ListViewItem itm = new ListViewItem(newArr);
                    this.Add2Quote3lv.Items.Add(itm);
                }
            }
        }
        private void populateCdb()
        {
            newBe.lists.RebuildQuoteDBLists();
            CommentDblv.Items.Clear();
            newBe.lists.QuoteDbList.Sort((x, y) => x.QcHeading.CompareTo(y.QcHeading));
            foreach (QuoteComment CmntQ in newBe.lists.QuoteDbList)
            {
                string[] newArr = new string[5];
                if (CmntQ.QcHeading.ToUpper().Contains(CDBsearchtb.Text.ToUpper()))
                {
                    newArr[0] = CmntQ.QcHeading;
                    newArr[1] = CmntQ.QcTextList.Count.ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    this.CommentDblv.Items.Add(itm);
                }
                else if (CDBsearchtb.Text == "")
                {
                    newArr[0] = CmntQ.QcHeading;
                    newArr[1] = CmntQ.QcTextList.Count.ToString();
                    ListViewItem itm = new ListViewItem(newArr);
                    this.CommentDblv.Items.Add(itm);
                }
            }
            
            NewCommentLv.Items.Clear();
            foreach (string Cmnt in NewCommentList)
            {
                NewCommentLv.Items.Add(Cmnt);
            }
        }
        private void PopulateGcom()
        {
            GcomLv.Items.Clear();
            //newBe.lists.GCommentdbList.Sort((x, y) => x.CompareTo(y));
            foreach (string cmt in newBe.lists.GCommentList)
            {
                GcomLv.Items.Add(cmt);
            }
        }
        private void PopulateGenCommentDb()
        {
            newBe.lists.RebuildGenQuoteDBLists();
            GCdbLv.Items.Clear();            
            newBe.lists.GCommentdbList.Sort((x, y) => x.CompareTo(y));
            foreach (string cmt in newBe.lists.GCommentdbList)
            {
                GCdbLv.Items.Add(cmt);
            }
        }
        private void AddNewCommbtn_Click(object sender, EventArgs e)
        {
            if (Newcommtb.Text != "")
            {
                NewCommentList.Add(this.Newcommtb.Text);
            }
            Newcommtb.Text = "";
            Populate();
        }

        private void QuoteClickbtn(object sender, EventArgs e)
        {
            this.errorProvider1.Clear();
            bool exitquick = false;
            if (newBe.lists.Add2QuoteList3.Count > 0 && Job3Headtb.Text == "")
            {
                this.errorProvider1.SetError(this.Job3Headtb, "cannot Generate quote if list is full and no Heading is present");
                exitquick = true;
            }
            if (newBe.lists.Add2QuoteList2.Count > 0 && Job2Headtb.Text == "")
            {
                this.errorProvider1.SetError(this.Job2Headtb, "cannot Generate quote if list is full and no Heading is present");
                exitquick = true;
            }
            if (newBe.lists.Add2QuoteList.Count > 0 && Job1Headtb.Text == "")
            {
                this.errorProvider1.SetError(this.Job1Headtb, "cannot Generate quote if list is full and no Heading is present");
                exitquick = true;
            }

            if (exitquick)
                return;

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
            s.Line.ForeColor.RGB = Color.FromArgb(81, 208, 146).ToArgb();
            s.TextFrame.TextRange.Text = "Quotation";
            s.TextFrame.TextRange.Font.Size = 24;
            s.TextFrame.TextRange.Font.Name = "Arial";
            s.TextFrame.TextRange.Font.Color = 0;

            //Footer:

            foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
            {
                Word.Range header = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                Word.Paragraph p88 = header.Paragraphs.Add(header);
                header.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                header.Text = "Version: 1.5";
                header.Font.Size = 8;
                header.Font.Bold = 0;
                // header.InsertParagraphAfter();
            }

            int spacing = 12;
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1a;
            oPara1a = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1a.Range.Text = "Quotation Date: \t\t" + QuoteDatetb.Text;
            oPara1a.Range.Font.Bold = 1;
            oPara1a.Range.Font.Size = 12;
            oPara1a.Range.Font.Name = "calibri";
            oPara1a.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1g;
            oPara1g = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1g.Range.Text = "Quote Number: \t\t\t" + Jobnumtb.Text;
            oPara1g.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1b;
            oPara1b = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1b.Range.Text = "Company: \t\t\t" + CmpNameTb.Text;
            oPara1b.Range.InsertParagraphAfter();


            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1c;
            oPara1c = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1c.Range.Text = "Email: \t\t\t\t" + CmpEmailtb.Text;
            oPara1c.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara1c.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            oPara1c.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1d;
            oPara1d = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1d.Range.Text = "Re: \t\t\t\t" + Projecttb.Text;
            oPara1d.Range.Font.Bold = 1;
            oPara1d.Range.Font.Size = 12;
            oPara1d.Range.Font.Name = "Calibri";
            //oPara1.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara1d.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara1d.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            oPara1d.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1e;
            oPara1e = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1e.Range.Text = "Location: \t\t\t" + Haddresstb.Text + ", " + Hsuburbtb.Text + " " + Hstatecb.Text + " " + Hpostcodetb.Text;
            oPara1e.Range.Font.Bold = 1;
            oPara1e.Range.Font.Size = 12;
            oPara1e.Range.Font.Name = "Calibri";
            //oPara1.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara1e.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara1e.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            oPara1e.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara2a;
            oPara2a = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara2a.Range.Text = "Hi " + ClientNametb.Text + ",";
            oPara2a.Range.Font.Bold = 0;
            oPara2a.Range.Font.Size = 12;
            oPara2a.Range.Font.Name = "Calibri";
            //oPara1.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara2a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara2a.Format.SpaceAfter = 10;    //20 pt spacing after paragraph.
            oPara2a.Range.InsertParagraphAfter();
            //Insert a paragraph at the beginning of the document.

            Word.Paragraph oPara2b;
            oPara2b = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara2b.Range.Text = "Thank you for your enquiry; we are pleased to provide our pricing for the supply " +
                "and installation of the below flooring systems as documented:";
            oPara2b.Range.Font.Bold = 0;
            oPara2b.Range.Font.Size = 12;
            oPara2b.Range.Font.Name = "Calibri";
            //oPara1.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara2b.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara2b.Format.SpaceAfter = 16;    //20 pt spacing after paragraph.
            oPara2b.Range.InsertParagraphAfter();

            Word.Paragraph oPara3;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara3.Range.Text = "Scope of Works";
            oPara3.Range.Font.Bold = 1;
            oPara3.Range.Font.Size = 22;
            oPara3.Range.Font.Name = "Calibri";
            oPara3.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Format.SpaceAfter = 16;    //20 pt spacing after paragraph.
            oPara3.Range.InsertParagraphAfter();

            //Word.Paragraph oPara4;
            //oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
            //if (Job1lmtb.Text == "")
            //{
            //    if (H1Warrant1Tb.Text == "")
            //        oPara4.Range.Text = "Allowance " + Job1m2tb.Text + "m2, and has a 5-year warranty";
            //    else
            //        oPara4.Range.Text = "Allowance " + Job1m2tb.Text + "m2, and has a " + H1Warrant1Tb.Text + "-year warranty";
            //    oPara4.Range.Font.Bold = 1;
            //    oPara4.Range.Font.Size = 14;
            //    oPara4.Range.Font.Name = "Calibri";
            //    oPara4.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            //    oPara4.Range.InsertParagraphAfter();
            //}
            //else
            //{
            //    if (H1Warrant1Tb.Text == "")
            //        oPara4.Range.Text = "Allowance " + Job1m2tb.Text + "m2, and has a 3-year warranty";
            //    else
            //        oPara4.Range.Text = "Allowance " + Job1m2tb.Text + "m2, and has a " + H1Warrant1Tb.Text + "-year warranty";
            //    oPara4.Range.Font.Bold = 1;
            //    oPara4.Range.Font.Size = 14;
            //    oPara4.Range.Font.Name = "Calibri";
            //    oPara4.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            //    oPara4.Range.InsertParagraphAfter();
            //}
            //oPara4.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            //oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //oPara4.Format.SpaceAfter = 16;    //20 pt spacing after paragraph.
            //oPara4.Range.InsertParagraphAfter();


            if (Job2cb.Checked || Job3cb.Checked)
                QuoteText(oDoc, newBe.lists.Add2QuoteList, Job1Headtb.Text, Job1m2tb.Text, Job1lmtb.Text, Job1Comptb.Text, J1Pricetb.Text, H1Warrant1Tb.Text, J1slipratetb.Text, true);
            else
                QuoteText(oDoc, newBe.lists.Add2QuoteList, Job1Headtb.Text, Job1m2tb.Text, Job1lmtb.Text, Job1Comptb.Text, J1Pricetb.Text, H1Warrant1Tb.Text, J1slipratetb.Text, false);

            if (Job3cb.Checked)
                QuoteText(oDoc, newBe.lists.Add2QuoteList2, Job2Headtb.Text, Job2m2tb.Text, Job2lmtb.Text, Job2Comptb.Text, J2Pricetb.Text, H1Warrant2Tb.Text, J2slipratetb.Text, true);
            else
                QuoteText(oDoc, newBe.lists.Add2QuoteList2, Job2Headtb.Text, Job2m2tb.Text, Job2lmtb.Text, Job2Comptb.Text, J2Pricetb.Text, H1Warrant2Tb.Text, J2slipratetb.Text, false);

            QuoteText(oDoc, newBe.lists.Add2QuoteList3, Job3Headtb.Text, Job3m2tb.Text, Job3lmtb.Text, Job3Comptb.Text, J3Pricetb.Text, H1Warrant3Tb.Text, J3slipratetb.Text, true);

            Word.Paragraph oPara5a;
            oPara5a = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara5a.Range.Text = "JOB PRICE: \t\t\t\t $" + Jpricetb.Text;
            oPara5a.Range.Font.Bold = 1;
            oPara5a.Range.Font.Size = 14;
            oPara5a.Range.Font.Name = "Calibri";
            oPara5a.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            oPara5a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara5a.Format.SpaceAfter = spacing / 2;    //20 pt spacing after paragraph.
            oPara5a.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara5b;
            oPara5b = oDoc.Content.Paragraphs.Add(ref oMissing);
            double Gst = 0;
            Gst = double.Parse(Jpricetb.Text) * 0.1;
            Gst = Math.Round(Gst, 2);
            oPara5b.Range.Text = "GST: \t\t\t\t\t $" + Gst.ToString();
            oPara5b.Range.Font.Bold = 1;
            oPara5b.Range.Font.Size = 14;
            oPara5b.Range.Font.Name = "Calibri";
            oPara5b.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            oPara5b.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara5b.Format.SpaceAfter = spacing / 2;    //20 pt spacing after paragraph.
            oPara5b.Range.InsertParagraphAfter();

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara5c;
            oPara5c = oDoc.Content.Paragraphs.Add(ref oMissing);
            double Total = double.Parse(Jpricetb.Text) + Gst;
            Total = Math.Round(Total, 2);
            oPara5c.Range.Text = "TOTAL PRICE INC GST: \t\t $" + Total.ToString();
            oPara5c.Range.Font.Bold = 1;
            oPara5c.Range.Font.Size = 14;
            oPara5c.Range.Font.Name = "Calibri";
            oPara5c.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            oPara5c.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara5c.Format.SpaceAfter = spacing + 10;    //20 pt spacing after paragraph.
            oPara5c.Range.InsertParagraphAfter();

            //Word.Range wrdRng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //object oPageBreak1 = Word.WdBreakType.wdPageBreak;
            //wrdRng1.InsertBreak(ref oPageBreak1);
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara6a = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara6a.Range.Text = "STANDARD TERMS & CONDITIONS:";
            oPara6a.Range.Font.Bold = 1;
            oPara6a.Range.Font.Size = 12;
            oPara6a.Range.Font.Name = "Calibri";
            oPara6a.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            oPara6a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara6a.Format.SpaceAfter = spacing;    //20 pt spacing after paragraph.
            oPara6a.Range.InsertParagraphAfter();

            Word.Paragraph oPara7 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara7.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            oPara7.Range.Font.Bold = 0;
            oPara7.Range.Text = "Payment:\t\t 30 days from completion of job";
            object oStart = oPara7.Range.Start;
            object oEnd = oPara7.Range.Start + 8;
            Word.Range rBold = oDoc.Range(ref oStart, ref oEnd);
            rBold.Bold = 1;
            oPara7.Range.InsertParagraphAfter();

            Word.Paragraph oPara7a = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara7a.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            oPara7a.Range.Font.Bold = 0;
            oPara7a.Range.Text = "Deposit:\t\t 20% on commencement - $" + Math.Round((Total * 0.2), 2).ToString();
            object oStart7a = oPara7a.Range.Start + 0;
            object oEnd7a = oPara7a.Range.Start + 8;
            Word.Range rBold7a = oDoc.Range(ref oStart7a, ref oEnd7a);
            rBold7a.Bold = 1;
            oPara7a.Range.InsertParagraphAfter();
            if (TnCcb.Checked == true)
            {
                Word.Paragraph oPara7b = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara7b.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                oPara7b.Range.Font.Bold = 0;
                oPara7b.Range.Text = "Warranty: \t\t " + Warrantytb.Text + "-year warranty";
                object oStart7b = oPara7b.Range.Start;
                object oEnd7b = oPara7b.Range.Start + 9;
                Word.Range rBold7b = oDoc.Range(ref oStart7b, ref oEnd7b);
                rBold7b.Bold = 1;
                oPara7b.Range.InsertParagraphAfter();

                Word.Paragraph oPara7c = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara7c.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                oPara7c.Range.Font.Bold = 0;
                oPara7c.Range.Text = "Completion: \t\t " + Completetb.Text + " consecutive days";
                object oStart7c = oPara7c.Range.Start;
                object oEnd7c = oPara7c.Range.Start + 11;
                Word.Range rBold7c = oDoc.Range(ref oStart7c, ref oEnd7c);
                rBold7c.Bold = 1;
                oPara7c.Format.SpaceAfter = spacing + 10;
                oPara7c.Range.InsertParagraphAfter();
            }
            GenerateGeneralComment();
            QuoteText(oDoc, GenCommentList, "General", "", "", "", "", "", "", false);

            foreach (Microsoft.Office.Interop.Word.Section wordSection in oDoc.Sections)
            {
                Word.Range footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //Word.Range footer = oDoc.ActiveWindow.Selection.Range;
                //Footer.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                Word.Paragraph p0 = footer.Paragraphs.Add(footer);
                footer.Text = " ";
                footer.InsertParagraphAfter();

                footer.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Paragraph p1 = footer.Paragraphs.Add(p0.Range);  
                if (Tatimcb.Checked)
                    footer.Text = "Tatim Hanna - Qepoxy Pty Ltd";
                if (Jakecb.Checked)
                    footer.Text = "Jacob Young - Qepoxy Pty Ltd";
                if (Jennacb.Checked)
                    footer.Text = "Jenna Hannam - Qepoxy Pty Ltd";
                footer.Font.Size = 8;
                footer.Font.Bold = 1;
                footer.InsertParagraphAfter();

                footer.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Paragraph p2 = footer.Paragraphs.Add(p1.Range);
                p2.Range.Text = "ABN: 33 628 159 428 | QBCC: 15140664 | PCCP: 251";
                p2.Range.Font.Size = 8;
                p2.Range.Font.Bold = 0;
                p2.Range.InsertParagraphAfter();

                Word.Paragraph p3 = footer.Paragraphs.Add(p2.Range);
                if (Tatimcb.Checked)
                {
                    p3.Range.Text = "tatim@qepoxy.com.au | Mobile: 0499 554 505 | Ph: 1800 737 699";
                    p3.Range.InsertParagraphAfter();
                }
                if (Jakecb.Checked)
                {
                    p3.Range.Text = "jacob@qepoxy.com.au | Mobile: 0447 071 047 | Ph: 1800 737 699";
                    p3.Range.InsertParagraphAfter();
                }
                if (Jennacb.Checked)
                {
                    p3.Range.Text = "jenna@qepoxy.com.au | Mobile: 0448 858 808 | Ph: 1800 737 699";
                    p3.Range.InsertParagraphAfter();
                }

                //footer.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Paragraph p4 = footer.Paragraphs.Add(p3.Range);
                p4.Range.Text = "www.Qepoxy.com.au";
                p4.Range.InsertParagraphAfter();

                Word.Paragraph p5 = footer.Paragraphs.Add(p4.Range);
                p5.Range.Text = " ";
                //p4.Range.InsertParagraphAfter();

                Word.Paragraph p6 = footer.Paragraphs.Add(p5.Range);                
                p6.Range.Fields.Add(p5.Range, Word.WdFieldType.wdFieldPage);
            }

            //oDoc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            //object nullobject = Missing.Value;
            //oDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            ////Formats the footer.
            //oDoc.ActiveWindow.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            //oDoc.ActiveWindow.Selection.Range.InsertParagraph();
            //oDoc.ActiveWindow.Selection.Text = Environment.NewLine;
            ////oDoc.ActiveWindow.Selection.Range.InsertAfter("text2");
            //oDoc.ActiveWindow.Selection.Font.Name = "Arial";
            //oDoc.ActiveWindow.Selection.Font.Size = 8;            
            ////Adds page numbers in the footer.    
            //Object CurrentPage = Word.WdFieldType.wdFieldPage; 
            //Object totalPages = Word.WdFieldType.wdFieldNumPages;

            //oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range, ref CurrentPage, ref nullobject, ref nullobject);
            //oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range, ref totalPages, ref nullobject, ref nullobject);

        }

        public void QuoteText(Word._Document oDoc, List<QuoteComment> inList, string HeadingText, string c_m2, string c_lm, string cdays, string jobPrice, string warranty,string SlipRating,bool Notes)
        {
            if (HeadingText != "" && HeadingText != "General")
            {
                object oMissing = System.Reflection.Missing.Value;
                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3a.Range.Text = HeadingText + ":";
                oPara3a.Range.Font.Bold = 1;
                oPara3a.Range.Font.Size = 12;
                oPara3a.Range.Font.Name = "Calibri";
                oPara3a.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                oPara3a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara3a.Format.SpaceAfter = 6;    //20 pt spacing after paragraph.
                oPara3a.Range.InsertParagraphAfter();

                foreach (QuoteComment Qc in inList)
                {
                    Word.Paragraph assets = oDoc.Content.Paragraphs.Add();
                    assets.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    assets.Range.Bold = 0;
                    assets.Range.Font.Size = 12;
                    assets.Range.ListFormat.ApplyBulletDefault();
                    //assets.Range.InsertParagraphAfter();

                    if (Qc.QcTextList.Count == 1)
                    {
                        assets.Range.InsertBefore(Qc.QcTextList[0]);
                    }
                    else
                    {
                        for (int i = 0; i < Qc.QcTextList.Count; i++)
                        {
                            if (i == Qc.QcTextList.Count - 1)
                            {
                                assets.Range.InsertBefore(Qc.QcTextList[i]);
                            }
                            else
                            {
                                assets.Range.InsertBefore(Qc.QcTextList[i] + "\n");
                            }
                        }
                    }
                }



                if (c_m2 != "" && c_lm == "")
                {
                    Word.Paragraph oPara4;
                    oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara4.Range.Text = "Allowance " + c_m2 + " m2";
                    oPara4.Range.Font.Bold = 1;
                    oPara4.Range.Font.Size = 12;
                    oPara4.Range.Font.Name = "Calibri";
                    oPara4.Format.SpaceAfter = 15;
                    oPara4.Range.InsertParagraphAfter();
                }
                else if (c_m2 != "" && c_lm != "")
                {
                    Word.Paragraph oPara4;
                    oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara4.Range.Text = "Allowance " + c_m2 + " m2, and " + c_lm + " lm";
                    oPara4.Range.Font.Bold = 1;
                    oPara4.Range.Font.Size = 12;
                    oPara4.Range.Font.Name = "Calibri";
                    oPara4.Format.SpaceAfter = 15;
                    oPara4.Range.InsertParagraphAfter();
                }

                if (jobPrice != "" || warranty != "" || cdays != "" || SlipRating != "")
                { 
                    Word.Paragraph oPara6a = oDoc.Content.Paragraphs.Add(ref oMissing);
                    if(Notes)
                    {
                        oPara6a.Range.Text = "Notes:";
                    }
                    else
                    {
                        oPara6a.Range.Text = "STANDARD TERMS & CONDITIONS:";
                    }
                    oPara6a.Range.Font.Bold = 1;
                    oPara6a.Range.Font.Size = 12;
                    oPara6a.Range.Font.Name = "Calibri";
                    oPara6a.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    oPara6a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara6a.Format.SpaceAfter = 10;
                    oPara6a.Range.InsertParagraphAfter();
                }

                if (jobPrice != "")
                {
                    Word.Paragraph oPara7a = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara7a.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    oPara7a.Range.Font.Bold = 0;
                    oPara7a.Range.Text = "Estimated Job Price: \t\t $ " + jobPrice;
                    object oStart7a = oPara7a.Range.Start + 0;
                    object oEnd7a = oPara7a.Range.Start + 20;
                    Word.Range rBold7a = oDoc.Range(ref oStart7a, ref oEnd7a);
                    rBold7a.Bold = 1;
                    //rBold7a.Format.SpaceAfter = 10;
                    oPara7a.Range.InsertParagraphAfter();
                }
                if (warranty != "")
                {
                    Word.Paragraph oPara7b = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara7b.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    oPara7b.Range.Font.Bold = 0;
                    oPara7b.Range.Text = "Warranty: \t\t " + warranty + "-year warranty";
                    object oStart7b = oPara7b.Range.Start;
                    object oEnd7b = oPara7b.Range.Start + 9;
                    Word.Range rBold7b = oDoc.Range(ref oStart7b, ref oEnd7b);
                    rBold7b.Bold = 1;
                    //rBold7b.Format.SpaceAfter = 10;
                    oPara7b.Range.InsertParagraphAfter();
                }
                if(SlipRating != "")
                {
                    Word.Paragraph oPara7c = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara7c.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    oPara7c.Range.Font.Bold = 0;
                    oPara7c.Range.Text = "SlipRating: \t\t " + SlipRating;
                    object oStart7c = oPara7c.Range.Start;
                    object oEnd7c = oPara7c.Range.Start + 11;
                    Word.Range rBold7c = oDoc.Range(ref oStart7c, ref oEnd7c);
                    rBold7c.Bold = 1;
                    //oPara7c.Format.SpaceAfter = 20 + 10;
                    oPara7c.Range.InsertParagraphAfter();

                }
                if(cdays != "")
                { 
                    Word.Paragraph oPara7c = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara7c.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    oPara7c.Range.Font.Bold = 0;
                    oPara7c.Range.Text = "Completion: \t\t " + cdays + " consecutive days";
                    object oStart7c = oPara7c.Range.Start;
                    object oEnd7c = oPara7c.Range.Start + 11;
                    Word.Range rBold7c = oDoc.Range(ref oStart7c, ref oEnd7c);
                    rBold7c.Bold = 1;
                    oPara7c.Format.SpaceAfter = 20 + 10;
                    oPara7c.Range.InsertParagraphAfter();
                }

            }
            else if (HeadingText == "General")
            {
                object oMissing = System.Reflection.Missing.Value;
                Word.Paragraph oPara3a;
                oPara3a = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara3a.Range.Text = HeadingText + ":";
                oPara3a.Range.Font.Bold = 1;
                oPara3a.Range.Font.Size = 14;
                oPara3a.Range.Font.Name = "Calibri";                
                oPara3a.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                oPara3a.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                oPara3a.Format.SpaceAfter = 6;    //20 pt spacing after paragraph.
                oPara3a.Range.InsertParagraphAfter();

                foreach (QuoteComment Qc in inList)
                {
                    Word.Paragraph assets = oDoc.Content.Paragraphs.Add();
                    assets.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    assets.Range.Bold = 0;
                    assets.Range.Font.Size = 12;
                    assets.Range.ListFormat.ApplyBulletDefault();
                    //assets.Range.InsertParagraphAfter();

                    if (Qc.QcTextList.Count == 1)
                    {
                        assets.Range.InsertBefore(Qc.QcTextList[0]);
                    }
                    else
                    {
                        for (int i = 0; i < Qc.QcTextList.Count; i++)
                        {
                            if (i == Qc.QcTextList.Count - 1)
                            {
                                assets.Range.InsertBefore(Qc.QcTextList[i]);
                            }
                            else
                            {
                                assets.Range.InsertBefore(Qc.QcTextList[i] + "\n");
                            }
                        }
                    }
                }
            }
        }
        private bool errorCheckName(string inputErr)
        {
            bool error = false;
            if (inputErr.Contains("'"))
            {
                error = true;
            }
            if (inputErr.Contains(","))
            {
                error = true;
            }
            if (inputErr.Contains(";"))
            {
                error = true;
            }
            if (inputErr.Contains(":"))
            {
                error = true;
            }

            return error;
        }
        private bool checktbError(TextBox tbin)
        {
            bool Error = false;
            //this.errorProvider1.Clear();
            if (errorCheckName(tbin.Text))
            {
                this.errorProvider1.SetError(tbin, "contains bad charecter either (, ; : ') ");
                Error = true;
            }
            return Error;
        }
        private bool checktbError(ComboBox tbin)
        {
            bool Error = false;
            //this.errorProvider1.Clear();
            if (errorCheckName(tbin.Text))
            {
                this.errorProvider1.SetError(tbin, "contains bad charecter , or ' ");
                Error = true;
            }
            return Error;
        }
        private void Save2DBbtn_Click(object sender, EventArgs e)
        {
            bool noError = true;
            this.errorProvider1.Clear();
            if (errorCheckName(Headingtb.Text))
            {
                this.errorProvider1.SetError(this.Headingtb, "contains bad charecter , or ' ");
                noError = false;
            }            
            if(noError)
            {
                QuoteComment newQc = new QuoteComment();
                newQc.QcHeading = Headingtb.Text;
                foreach (string str in NewCommentList)
                {
                    newQc.QcTextList.Add(str);
                }
                newBe.lists.AddtoCommentsDb(newQc);
                NewCommentList.Clear();
                Headingtb.Text = "";                
                populateCdb();
            }
        }

        private void EditDbitembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (CommentDblv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = CommentDblv.SelectedIndices[0];
                string item = CommentDblv.Items[Selected].Text;
                foreach (QuoteComment Qcitm in newBe.lists.QuoteDbList)
                {
                    if (Qcitm.QcHeading == item)
                    {
                        newQc = Qcitm;
                        break;
                    }
                }
                NewCommentList.Clear();
                foreach (string StrItm in newQc.QcTextList)
                {
                    NewCommentList.Add(StrItm);
                }
                Headingtb.Text = newQc.QcHeading;                
            }
            //populateCdb();
            populateCdb();
        }
        
        private void RmFrmDBbtn_Click(object sender, EventArgs e)
        {
            //int Selected = 0;
            //if (CommentDblv.SelectedIndices.Count > 0)
            //{
            //    QuoteComment newQc = new QuoteComment();
            //    Selected = CommentDblv.SelectedIndices[0];
            //    string item = CommentDblv.Items[Selected].Text;
            //    foreach (QuoteComment Qcitm in QuoteDbList)
            //    {
            //        if(item == Qcitm.QcHeading)
            //        {
            //            QuoteDbList.Remove(Qcitm);
            //            Populate();
            //            break;
            //        }
            //    }
            //}

            int Selected = 0;
            string item = "";
            if (CommentDblv.SelectedIndices.Count > 0)
            {
                Selected = CommentDblv.SelectedIndices[0];
                item = CommentDblv.Items[Selected].Text;               
            }
            if(item != "")
            {
                newBe.lists.DelItemFromCommentDb(item);
            }
            //Populate(); 
            populateCdb();
        }

        private void Add2Quotebtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            bool inDBlist = false;
            if (CommentDblv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = CommentDblv.SelectedIndices[0];
                string item = CommentDblv.Items[Selected].Text;

                foreach (QuoteComment Qucitm in newBe.lists.Add2QuoteList)
                {
                    if (Qucitm.QcHeading == item)
                    {
                        inDBlist = true;
                    }
                }
                if (!inDBlist)
                {
                    foreach (QuoteComment Qcitm in newBe.lists.QuoteDbList)
                    {
                        if (Qcitm.QcHeading == item)
                        {
                            newQc = Qcitm;
                            //foreach (string StrItm in Qcitm.QcTextList)
                            //{
                            //    newQc.QcTextList.Add(StrItm);
                            //}
                            break;
                        }
                    }

                    newBe.lists.Add2QuoteList.Add(newQc);
                    Populate();                    
                }
            }
        }   

        private void EditCommbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (NewCommentLv.SelectedIndices.Count > 0)
            {
                string newStr = "";
                Selected = NewCommentLv.SelectedIndices[0];
                string item = NewCommentLv.Items[Selected].Text;
                foreach (string Stritm in NewCommentList)
                {
                    if (Stritm == item)
                    {
                        newStr = Stritm;
                        Newcommtb.Text = Stritm;
                        NewCommentList.Remove(Stritm);
                        populateCdb();
                        break;
                    }
                }
            }
        }

        private void Upbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (NewCommentLv.SelectedIndices.Count > 0)
            {
                string newStr = "";
                Selected = NewCommentLv.SelectedIndices[0];
                string item = NewCommentLv.Items[Selected].Text;
                foreach (string Stritm in NewCommentList)
                {
                    if (Stritm == item)
                    {
                        newStr = item;
                        int index = NewCommentList.IndexOf(Stritm);
                        if (index == 0)
                        { }
                        else
                        {
                            NewCommentList.RemoveAt(index);
                            NewCommentList.Insert(index - 1, newStr);
                            Populate();
                            break;
                        }
                    }
                }
            }
        }

        private void Downbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (NewCommentLv.SelectedIndices.Count > 0)
            {
                string newStr = "";
                Selected = NewCommentLv.SelectedIndices[0];
                string item = NewCommentLv.Items[Selected].Text;
                foreach (string Stritm in NewCommentList)
                {
                    if (Stritm == item)
                    {
                        newStr = item;
                        int index = NewCommentList.IndexOf(Stritm);
                        if (index == NewCommentList.Count)
                        {

                        }
                        else
                        {
                            NewCommentList.RemoveAt(index);
                            NewCommentList.Insert(index + 1, newStr);
                            Populate();
                            break;
                        }
                    }
                }
            }

        }

        private void RemCmntFrmQuotebtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Add2Quote1lv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = Add2Quote1lv.SelectedIndices[0];
                newBe.lists.Add2QuoteList.RemoveAt(Selected);
            }
            Populate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string today = DateTime.Now.ToString("dd/MM/yyyy");
            this.QuoteDatetb.Text = today;
        }

        private void Tatimcb_CheckedChanged(object sender, EventArgs e)
        {
            if(Tatimcb.Checked)
            {
                Jakecb.Checked = false;
            }
        }

        private void Jakecb_CheckedChanged(object sender, EventArgs e)
        {
            if(Jakecb.Checked)
            {
                Tatimcb.Checked = false;
            }
        }

        private void Job2cb_CheckedChanged(object sender, EventArgs e)
        {
            if(Job2cb.Checked)
            {
                Job2Headtb.Enabled = true;
                Add2Quote2lv.Enabled = true;
                Add2Quote2btn.Enabled = true;
                RemCmntFrmQuote2btn.Enabled = true;
                Job2m2tb.Enabled = true;
                Job2lmtb.Enabled = true;
                Job2Comptb.Enabled = true;
                H1Warrant2Tb.Enabled = true;
                J2Pricetb.Enabled = true;
                J2slipratetb.Enabled = true;
            }
            else
            {
                Job2Headtb.Enabled = false;
                Add2Quote2lv.Enabled = false;
                Add2Quote2btn.Enabled = false;
                RemCmntFrmQuote2btn.Enabled = false;
                Job2m2tb.Enabled = false;
                Job2lmtb.Enabled = false;
                Job2Comptb.Enabled = false;
                H1Warrant2Tb.Enabled = false;
                J2Pricetb.Enabled = false;
                J2slipratetb.Enabled = false;

                Job2m2tb.Text = "";
                Job2lmtb.Text = "";
                Job2Comptb.Text = "";
                Job2Headtb.Text = "";
                J2slipratetb.Text = "";
                J2Pricetb.Text = "";
                H1Warrant2Tb.Text = "";

                if (newBe.lists.Add2QuoteList2.Count > 0)
                {
                    newBe.lists.Add2QuoteList2.Clear();
                }
                Populate();
            }
        }

        private void Job3cb_CheckedChanged(object sender, EventArgs e)
        {
            if (Job3cb.Checked)
            {
                Job3Headtb.Enabled = true;
                Add2Quote3lv.Enabled = true;
                Add2Quote3btn.Enabled = true;
                RemCmntFrmQuote3btn.Enabled = true;
                Job3m2tb.Enabled = true;
                Job3lmtb.Enabled = true;
                Job3Comptb.Enabled = true;
                H1Warrant3Tb.Enabled = true;
                J3Pricetb.Enabled = true;
                J3slipratetb.Enabled = true;
            }
            else
            {
                Job3Headtb.Enabled = false;
                Add2Quote3lv.Enabled = false;
                Add2Quote3btn.Enabled = false;
                RemCmntFrmQuote3btn.Enabled = false;
                Job3m2tb.Enabled = false;
                Job3lmtb.Enabled = false;
                Job3Comptb.Enabled = false;
                H1Warrant3Tb.Enabled = false;
                J3Pricetb.Enabled = false;
                J3slipratetb.Enabled = false;

                Job3m2tb.Text = "";
                Job3lmtb.Text = "";
                Job3Comptb.Text = "";
                Job3Headtb.Text = "";
                J3slipratetb.Text = "";
                J3Pricetb.Text = "";
                H1Warrant3Tb.Text = "";

                if (newBe.lists.Add2QuoteList3.Count > 0)
                {
                    newBe.lists.Add2QuoteList3.Clear();
                }
                Populate();
            }
        }

        private void Add2Quote2btn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            bool inDBlist = false;
            if (CommentDblv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = CommentDblv.SelectedIndices[0];
                string item = CommentDblv.Items[Selected].Text;

                foreach (QuoteComment Qucitm in newBe.lists.Add2QuoteList2)
                {
                    if (Qucitm.QcHeading == item)
                    {
                        inDBlist = true;
                    }
                }
                if (!inDBlist)
                {
                    foreach (QuoteComment Qcitm in newBe.lists.QuoteDbList)
                    {
                        if (Qcitm.QcHeading == item)
                        {
                            newQc = Qcitm;
                            //foreach (string StrItm in Qcitm.QcTextList)
                            //{
                            //    newQc.QcTextList.Add(StrItm);
                            //}
                            break;
                        }
                    }

                    newBe.lists.Add2QuoteList2.Add(newQc);
                    Populate();
                }
            }
        }

        private void RemCmntFrmQuote2btn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Add2Quote2lv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = Add2Quote2lv.SelectedIndices[0];
                newBe.lists.Add2QuoteList2.RemoveAt(Selected);
            }
            Populate();
        }

        private void Add2Quote3btn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            bool inDBlist = false;
            if (CommentDblv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = CommentDblv.SelectedIndices[0];
                string item = CommentDblv.Items[Selected].Text;

                foreach (QuoteComment Qucitm in newBe.lists.Add2QuoteList3)
                {
                    if (Qucitm.QcHeading == item)
                    {
                        inDBlist = true;
                    }
                }
                if (!inDBlist)
                {
                    foreach (QuoteComment Qcitm in newBe.lists.QuoteDbList)
                    {
                        if (Qcitm.QcHeading == item)
                        {
                            newQc = Qcitm;
                            //foreach (string StrItm in Qcitm.QcTextList)
                            //{
                            //    newQc.QcTextList.Add(StrItm);
                            //}
                            break;
                        }
                    }

                    newBe.lists.Add2QuoteList3.Add(newQc);
                    Populate();
                }
            }
        }

        private void RemCmntFrmQuote3btn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (Add2Quote3lv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = Add2Quote3lv.SelectedIndices[0];
                newBe.lists.Add2QuoteList3.RemoveAt(Selected);
            }
            Populate();
        }

        private void RemCommbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (NewCommentLv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = NewCommentLv.SelectedIndices[0];
                string item = NewCommentLv.Items[Selected].Text;
                foreach (string str in NewCommentList)
                {
                    if (item == str)
                    {
                        NewCommentList.Remove(str);
                        Populate();
                        break;
                    }
                }
            }
        }

        private void SaveCbtn_Click(object sender, EventArgs e)
        {
            bool oktosave = true;
            if (checktbError(Compnametb))
                oktosave = false;
            if (checktbError(Compemailtb))
                oktosave = false;
            if (checktbError(CompAddrtb))
                oktosave = false;
            if (checktbError(Suburbtb))
                oktosave = false;
            if (checktbError(CompStatecb))
                oktosave = false;
            if (checktbError(Postcodetb))
                oktosave = false;
            if (checktbError(ContactNametb))
                oktosave = false;



            if (oktosave)
            {
                Company newComp = new Company();
                newComp.CompName = Compnametb.Text;
                newComp.Email = Compemailtb.Text;
                newComp.Address = CompAddrtb.Text;
                newComp.Suburb = Suburbtb.Text;
                newComp.State = CompStatecb.Text;
                newComp.PostCode = Postcodetb.Text;
                newComp.ContactName = ContactNametb.Text;
                newComp.ContactNum = ContactNumtb.Text;

                updateBox("Saved to " + Compnametb.Text);
                newBe.lists.AddItemToCompanyList(newComp);
                newBe.lists.RebuildCompanyList();
                populateCompanyList();

                Compnametb.Text = "";
                Compemailtb.Text = "";
                CompAddrtb.Text = "";
                Suburbtb.Text = "";
                CompStatecb.Text = "";
                Postcodetb.Text = "";
                ContactNametb.Text = "";
                ContactNumtb.Text = "";
                tabControl3.SelectedTab = tabControl3.TabPages[1];

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Jobnum f4 = new Jobnum();
            f4.Owner = this;
            f4.Show();
        }

        private void Companycb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(CmpNameTb.Text != "")
            {
                foreach(Company cI in newBe.lists.Complist)
                {
                    if(CmpNameTb.Text == cI.CompName)
                    {
                        CmpEmailtb.Text = cI.Email;
                        Haddresstb.Text = cI.Address;
                        ClientNametb.Text = cI.CompName;
                    }
                }
            }
            else
            {
                CmpEmailtb.Text = "";
                Haddresstb.Text = "";
                ClientNametb.Text = "";
            }
        }

        private void EditCbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (CompanyLv.SelectedIndices.Count > 0)
            {
                Selected = CompanyLv.SelectedIndices[0];
                item = CompanyLv.Items[Selected].Text;                
            }

            if(item != "")
            {
                foreach(Company Citem in newBe.lists.Complist)
                { 
                    if(Citem.CompName == item)
                    { 
                        Compnametb.Text = Citem.CompName;
                        Compemailtb.Text = Citem.Email;
                        CompAddrtb.Text = Citem.Address;
                        Suburbtb.Text = Citem.Suburb;
                        CompStatecb.Text = Citem.State;
                        Postcodetb.Text = Citem.PostCode;
                        ContactNametb.Text = Citem.ContactName;
                        ContactNumtb.Text = Citem.ContactNum;
                    }
                }
            }
            tabControl3.SelectedTab = tabControl3.TabPages[2];
        }

        private void RemCbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (CompanyLv.SelectedIndices.Count > 0)
            {
                Selected = CompanyLv.SelectedIndices[0];
                item = CompanyLv.Items[Selected].Text;              
            }

            if(item != "")
            {
                newBe.lists.DelItemFromCompanyList(item);
                newBe.lists.RebuildCompanyList();
                populateCompanyList();
            }
        }

        private void CmpImportBtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (CompanyLv.SelectedIndices.Count > 0)
            {
                Selected = CompanyLv.SelectedIndices[0];
                item = CompanyLv.Items[Selected].Text;
            }

            foreach(Company Citem in newBe.lists.Complist)
            {
                if(item == Citem.CompName)
                { 
                    CmpNameTb.Text = Citem.CompName;
                    CmpEmailtb.Text = Citem.Email;
                    Haddresstb.Text = Citem.Address;
                    Hsuburbtb.Text = Citem.Suburb;
                    Hstatecb.Text = Citem.State;
                    Hpostcodetb.Text = Citem.PostCode;
                    ClientNametb.Text = Citem.ContactName;
                    ClientNumtb.Text = Citem.ContactNum;
                }
            }
            tabControl3.SelectedTab = tabControl3.TabPages[0];
        }

        private void GcdbSavebtn_Click(object sender, EventArgs e)
        {
            string item = GcNcTb.Text;
            if (item != "")
            {
                newBe.lists.AddtoGenComments(item);
                GcNcTb.Text = "";
                PopulateGenCommentDb();
            }
        }

        private void GcdbEditbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (GCdbLv.SelectedIndices.Count > 0)
            {
                Selected = GCdbLv.SelectedIndices[0];
                string item = GCdbLv.Items[Selected].Text;
                foreach (string GCmpitm in newBe.lists.GCommentdbList)
                {
                    if (item == GCmpitm)
                    {
                        GcNcTb.Text = GCmpitm;                        
                        PopulateGenCommentDb();
                        break;
                    }
                }
            }
        }

        private void GcdbRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (GCdbLv.SelectedIndices.Count > 0)
            {
                Selected = GCdbLv.SelectedIndices[0];
                item = GCdbLv.Items[Selected].Text; 
            }
            if(item != "")
            {
                newBe.lists.DelItemFromGenComms(item);
                PopulateGenCommentDb();
            }
        }

        private void GcAddbtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (GCdbLv.SelectedIndices.Count > 0)
            {
                Selected = GCdbLv.SelectedIndices[0];
                item = GCdbLv.Items[Selected].Text;
                newBe.lists.GCommentList.Add(item);
             }
            PopulateGcom();
        }

        private void GcRembtn_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            string item = "";
            if (GcomLv.SelectedIndices.Count > 0)
            {
                QuoteComment newQc = new QuoteComment();
                Selected = GcomLv.SelectedIndices[0];
                item = GCdbLv.Items[Selected].Text;
                newBe.lists.GCommentList.RemoveAt(Selected);                
            }
            PopulateGcom();
        }

        private void TnCcb_CheckedChanged(object sender, EventArgs e)
        {
            if (TnCcb.Checked)
            {
                Warrantytb.Enabled = true;
                Completetb.Enabled = true;
            }
            else
            {
                Warrantytb.Text = "";
                Completetb.Text = "";
                Warrantytb.Enabled = false;
                Completetb.Enabled = false;
            }
        }

        private void cmpClearBtn_Click(object sender, EventArgs e)
        {
            Jobnumtb.Text = "";
            QuoteDatetb.Text = "";
            CmpNameTb.Text = "";
            CmpEmailtb.Text = "";
            Haddresstb.Text = "";
            ClientNametb.Text = "";
            Projecttb.Text = "";
            Jpricetb.Text = "";
        }

        private void Qselectbtn_Click(object sender, EventArgs e)
        {
            
        }


        async public void QupdateBox(string Updatestr)
        {
            QuoteSaveLoadlbl.Text = Updatestr;
            await Task.Delay(5000);
            QuoteSaveLoadlbl.Text = "";
        }

        private void ClearFrmbtn_Click(object sender, EventArgs e)
        {
            Jobnumtb.Text = "";
            QuoteDatetb.Text = "";
            CmpNameTb.Text = "";
            CmpEmailtb.Text = "";
            Haddresstb.Text = "";
            Hsuburbtb.Text = "";
            Hstatecb.Text = "";
            Hpostcodetb.Text = "";
            ClientNumtb.Text = "";
            ClientNametb.Text = "";
            Projecttb.Text = "";
            Jpricetb.Text = "";

            Job1m2tb.Text = "";
            Job1lmtb.Text = "";
            Job1Comptb.Text = "";
            Job1Headtb.Text = "";
            J1slipratetb.Text = "";
            J1Pricetb.Text = "";
            H1Warrant1Tb.Text = "";

            Job2cb.Checked = false;
            Job2m2tb.Text = "";
            Job2lmtb.Text = "";
            Job2Comptb.Text = "";
            Job2Headtb.Text = "";
            J2slipratetb.Text = "";
            J2Pricetb.Text = "";
            H1Warrant2Tb.Text = "";

            Job3cb.Checked = false;
            Job3m2tb.Text = "";
            Job3lmtb.Text = "";
            Job3Comptb.Text = "";
            Job3Headtb.Text = "";
            J3slipratetb.Text = "";
            J3Pricetb.Text = "";
            H1Warrant3Tb.Text = "";

            Rubbcb.Checked = false;
            PH3cb.Checked = false;
            DryTimecb.Checked = false;
            Jointcb.Checked = false;
            FloorColourcb.Checked = false;
            SlipRcb.Checked = false;
            QuoteStagecb.Checked = false;
            Aheadcb.Checked = false;
            ContJcb.Checked = false;

            JCCtb.Text = "";
            Coltb.Text = "";
            SRtb.Text = "";
            QuoteStagestb.Text = "";
            Aheadbox.Text = "";
            ContJtb.Text = "";
            QuoteNametb.Text = "";

            newBe.lists.Add2QuoteList.Clear();
            newBe.lists.Add2QuoteList2.Clear();
            newBe.lists.Add2QuoteList3.Clear();
            newBe.lists.GCommentList.Clear();
            PopulateGcom();
            Populate();
        }

        private void CDBsearchbtn_Click(object sender, EventArgs e)
        {
            
        }

        private void QexportDatabtn_Click(object sender, EventArgs e)
        {
            bool inlist = false;
            foreach (Quote_Block qb in newBe.lists.Q_bList)
            {
                if(QuoteNametb.Text == qb.QName)
                {
                    inlist = true;
                }
            }

            if(inlist)
            {
                QuoteBlock qbl = new QuoteBlock();
                bool oktosave = true;
                this.errorProvider1.Clear();

                if (checktbError(QuoteNametb))
                    oktosave = false;
                if (checktbError(Jobnumtb))
                    oktosave = false;
                if (checktbError(QuoteDatetb))
                    oktosave = false;
                if (checktbError(CmpNameTb))
                    oktosave = false;
                if (checktbError(CmpEmailtb))
                    oktosave = false;
                if (checktbError(Haddresstb))
                    oktosave = false;
                if (checktbError(Hsuburbtb))
                    oktosave = false;
                if (checktbError(Hpostcodetb))
                    oktosave = false;
                if (checktbError(ClientNametb))
                    oktosave = false;
                if (checktbError(ClientNumtb))
                    oktosave = false;
                if (checktbError(Projecttb))
                    oktosave = false;
                if (checktbError(Jpricetb))
                    oktosave = false;
                if (checktbError(Job1m2tb))
                    oktosave = false;
                if (checktbError(Job1lmtb))
                    oktosave = false;
                if (checktbError(Job1Comptb))
                    oktosave = false;
                if (checktbError(Job1Headtb))
                    oktosave = false;
                if (checktbError(J1slipratetb))
                    oktosave = false;
                if (checktbError(J1Pricetb))
                    oktosave = false;
                if (checktbError(H1Warrant1Tb))
                    oktosave = false;
                if (checktbError(Job2m2tb))
                    oktosave = false;
                if (checktbError(Job2lmtb))
                    oktosave = false;
                if (checktbError(Job2Comptb))
                    oktosave = false;
                if (checktbError(J2slipratetb))
                    oktosave = false;
                if (checktbError(Job2Headtb))
                    oktosave = false;
                if (checktbError(J2Pricetb))
                    oktosave = false;
                if (checktbError(H1Warrant2Tb))
                    oktosave = false;
                if (checktbError(Job3m2tb))
                    oktosave = false;
                if (checktbError(Job3Comptb))
                    oktosave = false;
                if (checktbError(Job3lmtb))
                    oktosave = false;
                if (checktbError(Job3Headtb))
                    oktosave = false;
                if (checktbError(J3slipratetb))
                    oktosave = false;
                if (checktbError(J3Pricetb))
                    oktosave = false;
                if (checktbError(H1Warrant3Tb))
                    oktosave = false;
                if (checktbError(JCCtb))
                    oktosave = false;
                if (checktbError(Coltb))
                    oktosave = false;
                if (checktbError(SRtb))
                    oktosave = false;
                if (checktbError(QuoteStagestb))
                    oktosave = false;
                if (checktbError(QuoteStagestb))
                    oktosave = false;
                if (checktbError(ContJtb))
                    oktosave = false;

                if (oktosave)
                {
                    qbl.QName = QuoteNametb.Text;
                    qbl.Qinfo.Jobnum = Jobnumtb.Text;
                    qbl.Qinfo.JobDate = QuoteDatetb.Text;
                    qbl.Qinfo.CompanyName = CmpNameTb.Text;
                    qbl.Qinfo.Email = CmpEmailtb.Text;
                    qbl.Qinfo.Address = Haddresstb.Text;
                    qbl.Qinfo.Suburb = Hsuburbtb.Text;
                    qbl.Qinfo.State = Hstatecb.Text;
                    qbl.Qinfo.Postcode = Hpostcodetb.Text;
                    qbl.Qinfo.ClientName = ClientNametb.Text;
                    qbl.Qinfo.ClientNum = ClientNumtb.Text;
                    qbl.Qinfo.Project = Projecttb.Text;
                    qbl.Qinfo.TotalEst = Jpricetb.Text;
                    qbl.Qinfo.Q1Allowance = Job1m2tb.Text;
                    qbl.Qinfo.Q1Lm = Job1lmtb.Text;
                    qbl.Qinfo.Q1Days = Job1Comptb.Text;
                    qbl.Qinfo.Q1Heading = Job1Headtb.Text;
                    qbl.Qinfo.Q1SlipRate = J1slipratetb.Text;
                    qbl.Qinfo.Q1StageVal = J1Pricetb.Text;
                    qbl.Qinfo.Q1Warrenty = H1Warrant1Tb.Text;
                    qbl.Qinfo.Q2Enable = Job2cb.Checked;
                    qbl.Qinfo.Q2Allowance = Job2m2tb.Text;
                    qbl.Qinfo.Q2Lm = Job2lmtb.Text;
                    qbl.Qinfo.Q2Days = Job2Comptb.Text;
                    qbl.Qinfo.Q2Heading = Job2Headtb.Text;
                    qbl.Qinfo.Q2SlipRate = J2slipratetb.Text;
                    qbl.Qinfo.Q2StageVal = J2Pricetb.Text;
                    qbl.Qinfo.Q2Warrenty = H1Warrant2Tb.Text;
                    qbl.Qinfo.Q3Enable = Job3cb.Checked;
                    qbl.Qinfo.Q3Allowance = Job3m2tb.Text;
                    qbl.Qinfo.Q3Lm = Job3lmtb.Text;
                    qbl.Qinfo.Q3Days = Job3Comptb.Text;
                    qbl.Qinfo.Q3Heading = Job3Headtb.Text;
                    qbl.Qinfo.Q3SlipRate = J3slipratetb.Text;
                    qbl.Qinfo.Q3StageVal = J3Pricetb.Text;
                    qbl.Qinfo.Q3Warrenty = H1Warrant3Tb.Text;
                    qbl.Qinfo.Rubbish = Rubbcb.Checked;
                    qbl.Qinfo.Phase3 = PH3cb.Checked;
                    qbl.Qinfo.Drytime = DryTimecb.Checked;
                    qbl.Qinfo.JointColour = Jointcb.Checked;
                    qbl.Qinfo.Colour = FloorColourcb.Checked;
                    qbl.Qinfo.SlipRes = SlipRcb.Checked;
                    qbl.Qinfo.QuoteStages = QuoteStagecb.Checked;
                    qbl.Qinfo.ArrowHead = Aheadcb.Checked;
                    qbl.Qinfo.AllowCtrlJoins = ContJcb.Checked;
                    qbl.Qinfo.JointCstr = JCCtb.Text;
                    qbl.Qinfo.Colstr = Coltb.Text;
                    qbl.Qinfo.SlipRstr = SRtb.Text;
                    qbl.Qinfo.QuoteStagesstr = QuoteStagestb.Text;
                    qbl.Qinfo.Arrow = Aheadbox.Text;
                    qbl.Qinfo.AllowCtrlstr = ContJtb.Text;

                    foreach (QuoteComment qcom in newBe.lists.Add2QuoteList)
                        qbl.Qinfo.Q1CommentList.Add(qcom);
                    foreach (QuoteComment qcom in newBe.lists.Add2QuoteList2)
                        qbl.Qinfo.Q2CommentList.Add(qcom);
                    foreach (QuoteComment qcom in newBe.lists.Add2QuoteList3)
                        qbl.Qinfo.Q3CommentList.Add(qcom);
                    foreach (string qcom in newBe.lists.GCommentList)
                        qbl.Qinfo.CommentList.Add(qcom);

                    newBe.lists.AddtoQuoteGenLists(qbl);
                }
                else
                {
                    this.errorProvider1.SetError(QexportDatabtn, "Can't Save due to errors");
                }
            }            
        }

        private void QimportDatabrn_Click(object sender, EventArgs e)
        {
            string quotename = QuoteNametb.Text;
            newBe.lists.RebuildQuoteList();
            foreach(Quote_Block qblock in newBe.lists.Q_bList)
            {
                if(quotename == qblock.QName)
                {
                    Jobnumtb.Text = qblock.Qinfo.Jobnum;
                    QuoteDatetb.Text = qblock.Qinfo.JobDate;
                    CmpNameTb.Text = qblock.Qinfo.CompanyName;
                    CmpEmailtb.Text = qblock.Qinfo.Email;

                    Haddresstb.Text = qblock.Qinfo.Address;
                    Hsuburbtb.Text = qblock.Qinfo.Suburb;
                    Hstatecb.Text = qblock.Qinfo.State;
                    Hpostcodetb.Text = qblock.Qinfo.Postcode;

                    ClientNametb.Text = qblock.Qinfo.ClientName;
                    ClientNumtb.Text = qblock.Qinfo.ClientNum;

                    Projecttb.Text = qblock.Qinfo.Project;
                    Jpricetb.Text = qblock.Qinfo.TotalEst;

                    Job1m2tb.Text = qblock.Qinfo.Q1Allowance;
                    Job1lmtb.Text = qblock.Qinfo.Q1Lm;
                    Job1Comptb.Text = qblock.Qinfo.Q1Days;
                    Job1Headtb.Text = qblock.Qinfo.Q1Heading;
                    J1slipratetb.Text = qblock.Qinfo.Q1SlipRate;
                    J1Pricetb.Text = qblock.Qinfo.Q1StageVal;
                    H1Warrant1Tb.Text = qblock.Qinfo.Q1Warrenty;

                    Job2cb.Checked = qblock.Qinfo.Q2Enable;
                    Job2m2tb.Text = qblock.Qinfo.Q2Allowance;
                    Job2lmtb.Text = qblock.Qinfo.Q2Lm;
                    Job2Comptb.Text = qblock.Qinfo.Q2Days;
                    Job2Headtb.Text = qblock.Qinfo.Q2Heading;
                    J2slipratetb.Text = qblock.Qinfo.Q2SlipRate;
                    J2Pricetb.Text = qblock.Qinfo.Q2StageVal;
                    H1Warrant2Tb.Text = qblock.Qinfo.Q2Warrenty;

                    Job3cb.Checked = qblock.Qinfo.Q3Enable;
                    Job3m2tb.Text = qblock.Qinfo.Q3Allowance;
                    Job3lmtb.Text = qblock.Qinfo.Q3Lm;
                    Job3Comptb.Text = qblock.Qinfo.Q3Days;
                    Job3Headtb.Text = qblock.Qinfo.Q3Heading;
                    J3slipratetb.Text = qblock.Qinfo.Q3SlipRate;
                    J3Pricetb.Text = qblock.Qinfo.Q3StageVal;
                    H1Warrant3Tb.Text = qblock.Qinfo.Q3Warrenty;

                    Rubbcb.Checked = qblock.Qinfo.Rubbish;
                    PH3cb.Checked = qblock.Qinfo.Phase3;
                    DryTimecb.Checked = qblock.Qinfo.Drytime;
                    Jointcb.Checked = qblock.Qinfo.JointColour;
                    FloorColourcb.Checked = qblock.Qinfo.Colour;
                    SlipRcb.Checked = qblock.Qinfo.SlipRes;
                    QuoteStagecb.Checked = qblock.Qinfo.QuoteStages;
                    Aheadcb.Checked = qblock.Qinfo.ArrowHead;
                    ContJcb.Checked = qblock.Qinfo.AllowCtrlJoins;

                    JCCtb.Text = qblock.Qinfo.JointCstr;
                    Coltb.Text = qblock.Qinfo.Colstr;
                    SRtb.Text = qblock.Qinfo.SlipRstr;
                    QuoteStagestb.Text = qblock.Qinfo.QuoteStagesstr;
                    Aheadbox.Text = qblock.Qinfo.Arrow;
                    ContJtb.Text = qblock.Qinfo.AllowCtrlstr;
                    //QuoteNametb.Text = qblock.Qinfo.q;

                    newBe.lists.Add2QuoteList.Clear();
                    newBe.lists.Add2QuoteList2.Clear();
                    newBe.lists.Add2QuoteList3.Clear();
                    newBe.lists.GCommentList.Clear();

                    foreach (QuoteComment qcom in qblock.Qinfo.Q1CommentList)
                        newBe.lists.Add2QuoteList.Add(qcom);
                    foreach (QuoteComment qcom in qblock.Qinfo.Q2CommentList)
                        newBe.lists.Add2QuoteList2.Add(qcom);
                    foreach (QuoteComment qcom in qblock.Qinfo.Q3CommentList)
                        newBe.lists.Add2QuoteList3.Add(qcom);
                    foreach (string qcom in qblock.Qinfo.CommentList)
                        newBe.lists.GCommentList.Add(qcom);

                    //PopulateGenCommentDb();
                    PopulateGcom();
                    Populate();
                    populateCdb();

                }
            }
        }

        private void SearchTb_TextChanged(object sender, EventArgs e)
        {
            Populate();
        }

        private void CDBsearchtb_TextChanged(object sender, EventArgs e)
        {
            populateCdb();
        }

        private void QuotelistLv_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            loadQuotes();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            int Selected = 0;
            if (GCdbLv.SelectedIndices.Count > 0)
            { 
                Selected = GCdbLv.SelectedIndices[0];
                string Qitem = newBe.lists.GCommentdbList[Selected];

                foreach(string Qcitm in newBe.lists.GCommentdbList)
                {
                    if (Qcitm == Qitem)
                    {

                        int index = newBe.lists.GCommentdbList.IndexOf(Qitem);
                        if (index <= 0)
                        { GCdbLv.Items[index].Selected = true; }
                        else
                        {
                            newBe.lists.GCommentdbList.RemoveAt(index);
                            newBe.lists.GCommentdbList.Insert(index - 1, Qitem);
                            PopulateGenCommentDb();
                            GCdbLv.Items[index - 1].Selected = true;
                            break;
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int Selected = 0;
            if (GCdbLv.SelectedIndices.Count > 0)
            {

                Selected = GCdbLv.SelectedIndices[0];
                string Qitem = newBe.lists.GCommentdbList[Selected];

                foreach (string Qcitm in newBe.lists.GCommentdbList)
                {
                    if (Qcitm == Qitem)
                    {
                        int index = newBe.lists.GCommentdbList.IndexOf(Qitem);
                        if ((index + 1) == GCdbLv.Items.Count)
                        {
                            GCdbLv.Items[index].Selected = true;
                        }
                        else
                        {
                            newBe.lists.GCommentdbList.RemoveAt(index);
                            newBe.lists.GCommentdbList.Insert(index + 1, Qitem);
                            PopulateGenCommentDb();
                            GCdbLv.Items[index + 1].Selected = true;
                            index = index + 1;
                            break;
                        }
                    }
                }
            }
        }

        private void ClearCompBtn_Click(object sender, EventArgs e)
        {
            Compnametb.Text = "";
            Compemailtb.Text = "";
            CompAddrtb.Text = "";
            Suburbtb.Text = "";
            CompStatecb.Text = "";
            Postcodetb.Text = "";
            ContactNametb.Text = "";
            ContactNumtb.Text = "";            
        }
        async public void updateBox(string Updatestr)
        {
            StatusLabel1.Text = Updatestr;
            await Task.Delay(5000);
            StatusLabel1.Text = "";
        }

    }
}
