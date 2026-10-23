using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Qepoxy
{
    public partial class DataControl : Form
    {
        Backend newBackend;
        public DataControl(Backend backend)
        {
            newBackend = backend;
            InitializeComponent();
            this.SQLerrorlv.Scrollable = true;
            this.SQLerrorlv.View = View.Details;
            this.SQLerrorlv.GridLines = false;
            this.SQLerrorlv.FullRowSelect = true;
            this.SQLerrorlv.Columns.Add("Data Errors", 1500);

            DataPG.Value = 0;
            DataPG.Maximum = 100;

        }
        //public async void updateListview()
        public void updateListview()
        {
            this.SQLerrorlv.Items.Clear();
            if (newBackend.lists.SQLErrorList.Count > 0)
            {
                foreach (string str in newBackend.lists.SQLErrorList)
                {
                    this.SQLerrorlv.Items.Add(str);
                }
            }
            //await Task.Delay(5000);
            //updateListview();

        }
        private async void button3_Click(object sender, EventArgs e)
        {
            if(DataLoadtb.Text != "")
            {
                DataPG.Value = 0;
                newBackend.lists.SQLErrorList.Add("%% Dropping Tables %%");
                updateListview();
                await Task.Run(() => newBackend.lists.DropAllTables());
                newBackend.lists.SQLErrorList.Add("%%-- Finished Dropping Tables --%%");
                DataPG.Value += 100/6;
                newBackend.lists.SQLErrorList.Add("%% Importing Product, Packs, Systems, hrrates and Additions %%"); 
                updateListview();

                await Task.Run(() => newBackend.lists.ImportAllDataFromDiskToSQL(DataLoadtb.Text));
                DataPG.Value += 100 / 6;
                newBackend.lists.SQLErrorList.Add("%%-- Finished Product, Packs and Systems, hrrates and Additions --%%");
                newBackend.lists.SQLErrorList.Add("%% Importing Quotes &&");
                updateListview();


                await Task.Run(() => newBackend.lists.ImportQuotesFromDiskToSQL(DataLoadtb.Text));
                DataPG.Value += 100 / 6;
                newBackend.lists.SQLErrorList.Add("%%-- Finished Importing Quotes --%%");
                newBackend.lists.SQLErrorList.Add("%% Importing CompanyData %%");
                updateListview();
                
                await Task.Run(() => newBackend.lists.ImportCompanyDataToSql(DataLoadtb.Text));
                DataPG.Value += 100 / 6;
                newBackend.lists.SQLErrorList.Add("%%-- Finished Importing CompanyData --%%");
                newBackend.lists.SQLErrorList.Add("%% Importing Comments DB %%");
                updateListview();

                await Task.Run(() => newBackend.lists.ImportCommentDbToSql(DataLoadtb.Text));
                DataPG.Value += 100 / 6;
                newBackend.lists.SQLErrorList.Add("%%-- Finished Importing Comments --%%"); 
                newBackend.lists.SQLErrorList.Add("%% Rebuilding all lists %%");
                updateListview();

                await Task.Run(() => newBackend.lists.LoadSQLData());
                DataPG.Value += 100 / 6;
                newBackend.lists.SQLErrorList.Add("%%-- Finished Rebuilding all lists --%%");
                newBackend.lists.SQLErrorList.Add(" ");
                newBackend.lists.SQLErrorList.Add("%%-- Safe to close the Importer --%% ");
                updateListview();

                DataPG.Value = 100;
            }           
        }

        private void button6_Click(object sender, EventArgs e)
        {


            DataPG.Value = 0;
            newBackend.lists.SQLErrorList.Clear();
            newBackend.lists.SQLErrorList.Add("%% Saving Systems Data %%");
            newBackend.lists.DataSaveSystemsData(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving Company Data %%");
            newBackend.lists.dataSaveCompanyData(Datasavelocationtb.Text);
            updateListview();

            newBackend.lists.SQLErrorList.Add("%% Saving Pack Data %%");
            newBackend.lists.DataSavePackData(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving PricedItems Data %%");
            newBackend.lists.DataSavePriceItemsData(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving Hrrates Data %%");
            newBackend.lists.DataSaveHrRates(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving Additions Data %%");
            newBackend.lists.DataSaveAdditions(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving QuoteFileSystem Data %%");
            newBackend.lists.DataSaveQuoteFSdata(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving Quotes Data %%");
            newBackend.lists.DatasaveQuoteData(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving General Comments Data %%");
            newBackend.lists.DataSaveGCommentsDb(Datasavelocationtb.Text);
            updateListview();
            DataPG.Value += 100 / 8;

            newBackend.lists.SQLErrorList.Add("%% Saving QuoteDatabase %%");
            newBackend.lists.DataSaveDbToDisk(Datasavelocationtb.Text);
            

            newBackend.lists.SQLErrorList.Add(" ");
            newBackend.lists.SQLErrorList.Add("%%-- Finished Saving data lists --%%");
            newBackend.lists.SQLErrorList.Add(" ");
            newBackend.lists.SQLErrorList.Add("%%-- Safe to close --%% ");
            updateListview();
            DataPG.Value = 100;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                MessageBox.Show("You selected: " + dialog.FileName);
                DataLoadtb.Text = dialog.FileName+"\\";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                MessageBox.Show("You selected: " + dialog.FileName);
                Datasavelocationtb.Text = dialog.FileName + "\\";
            }
        }

        private void Clearbtn_Click(object sender, EventArgs e)
        {
            this.SQLerrorlv.Items.Clear();
        }

        private void Legacycb_CheckedChanged(object sender, EventArgs e)
        {
            if (Legacycb.Checked == true)
            {
                newBackend.lists.legacyData = true;
            }
            else if(Legacycb.Checked == false)
            {
                newBackend.lists.legacyData = false;
            }
        }
    }
}
