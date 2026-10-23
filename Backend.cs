using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using MySql.Data.MySqlClient;
using System.Data;


namespace Qepoxy
{
    public class Backend
    {
        //double sqrmeter;
        //double perimeter;
        public dataLists lists = new dataLists();
        public bool Startup = true;
        public bool GobalDC = false;
        public Backend()
        {
            lists.LoadLists();
        }
    }

    public class dataLists
    {   
        //-------------------------------------------------------------------
        // main window lists
        //-------------------------------------------------------------------
        public List<Pack> PackList = new List<Pack>();
        public List<ProductSystem> SystemsList = new List<ProductSystem>();
        public List<Priceitem> PriceList = new List<Priceitem>();
        public HourRates HrRates = new HourRates();
        public List<Costing> POCostList = new List<Costing>();
        public List<CSstage> COstageList = new List<CSstage>();
        //-------------------------------------------------------------------
        // quote window lists
        //-------------------------------------------------------------------
        public List<Company> Complist = new List<Company>();
        public List<QuoteBlock> QbList = new List<QuoteBlock>();
        public List<Quote_Block> Q_bList = new List<Quote_Block>();
        public List<QuoteFS> QuoteFSdata = new List<QuoteFS>();
        public List<QuoteComment> Add2QuoteList = new List<QuoteComment>();
        public List<QuoteComment> Add2QuoteList2 = new List<QuoteComment>();
        public List<QuoteComment> Add2QuoteList3 = new List<QuoteComment>();
        public List<String> GCommentList = new List<String>();
        public List<Hrrates> HrrateList = new List<Hrrates>();
        public List<HrCombo> HrComboList = new List<HrCombo>();
        public List<string> AdditionsList = new List<string>();
        public List<Adds> AddsList = new List<Adds>();

        public List<QuoteComment> QuoteDbList = new List<QuoteComment>();
        public List<String> GCommentdbList = new List<String>();

        public int itemsChangedCount = 0;
        public int SystemItemsAdded = 0;
        public int SystemItemsRemoved = 0;
        public string ProgressBar1;
        //StreamWriter logWriter = new StreamWriter(@"..\systemLog.log");
        //-------------------------------------------------------------------
        // sql Database Vars
        //-------------------------------------------------------------------
        MySqlConnection sqlConn = new MySqlConnection();
        MySqlCommand sqlCmd = new MySqlCommand();
        string sqlQuery;
        MySqlDataAdapter Dta = new MySqlDataAdapter();
        MySqlDataReader sqlRd;
        public List<string> SQLErrorList = new List<string>();
        //public string server = "192.168.56.1";
        //public string server = "192.168.0.10";
        //public string server = "114.75.162.220";
        //public string server = "1.146.49.114";
        public string server = "qepoxyserver.ddns.net";
        public string username = "Admin";
        public string password = "SqlAdmin";
        string database = "qepoxy2";
        public bool legacyData = false;
        public bool serverResult = false;
        static int ID = 1; 
        static int ID2 = 1;
        static int ID3 = 1;

        public void LoadLists()
        {
            //string user1 = "Tatim@qepoxy";
            //string pw1 = "password";
            sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database + ";";
            ForceCreateTables();
        }

        private void ForceCreateTables()
        {
            
            createTables("create table IF not exists qepoxy2.quote_blockdata(Qname varchar(30), Type varchar(30) ,PackId int, InputName varchar(100), Ammount varchar(100), Area varchar(30), Lm varchar(30), Cost varchar(30),PackAmt varchar(30)," +
                       "PackSystemName varchar(150),Packvol varchar(30), Price varchar(30), Stage varchar(30), Symbol varchar(30), Systems varchar(150));");
            createTables("create table IF not exists qepoxy2.quote_hrsdata(Qname varchar(100), Type varchar(30) ,PackId int, Hours varchar(10), Lm varchar(100), Area varchar(50), Method varchar(150), Stage varchar(150),Systems varchar(150),Techhrs varchar(30));");
            createTables("create table IF not exists qepoxy2.quote_blockadds(Qname varchar(30), Type varchar(30) ,PackId int, ItemName varchar(50), Descr varchar(100), Ammount varchar(50), Cost varchar(30));");
            createTables("create table IF not exists qepoxy2.additions(ItemName varchar(50));");
            createTables("create table IF not exists qepoxy2.companydata(CName varchar(100), EmailAddr varchar(50), Address varchar(50), Suburb varchar(50), State varchar(10), Postcode varchar(10), ContactName varchar(50), ContactNum varchar(50)); ");
            createTables("create table IF not exists qepoxy2.Gencomments(Comment varchar(200));");
            createTables("create table IF not exists qepoxy2.CommentsDB( Heading varchar(200), subHeading char(20), Comment varchar(300));");
            
            createTables("create table IF not exists qepoxy2.quoteoptions(QName varchar(100), jobnum char(30), jobDate char(30), companyName char(30),email char(30),Address char(100), " +
                    "Suburb char(30),State char(30), Postcode char(30), clientName char(30),clientNum char(30),project char(100),totalEst char(30)," +
                                      "Q1Allowance char(30),Q1Lm char(30),Q1Days char(30),Q1Heading char(100),Q1SlipRate char(30),Q1StageVal char(30),Q1Warrenty char(30)," +
                    "Q2Enable varchar(20), Q2Allowance char(30), Q2Lm char(30), Q2Days char(30), Q2Heading char(100), Q2SlipRate char(30), Q2StageVal char(30), Q2Warrenty char(30)," +
                    "Q3Enable varchar(20), Q3Allowance char(30), Q3Lm char(30), Q3Days char(30), Q3Heading char(100), Q3SlipRate char(30), Q3StageVal char(30), Q3Warrenty char(30)," +
                    "Rubbish varchar(20), Phase3 varchar(20), Drytime varchar(20), JointColour varchar(20), Colour varchar(20), SlipRes varchar(20), QuoteStages varchar(20), ArrowHead varchar(20), AllowCtrlJoins varchar(20)," +
                    "jointCstr varchar(30), Colourstr varchar(30), slipR varchar(30), QuoteStagesStr varchar(30), ArrowH varchar(30), AllowCtrl varchar(30));");

            createTables("create table IF not exists qepoxy2.hrratesettings(HrName varchar(100), HrWeight char(20), primary key(HrName), unique(HrName));");
            createTables("create table IF not exists qepoxy2.FSdata(FsName varchar(100), Type char(20), primary key(FsName), unique(FsName));");
            createTables("create table IF not exists qepoxy2.PackData(PaName varchar(100), Price double, Weight double, Volume double,Coverage double,Unit char(15),primary key(PaName), unique(PaName));");
            createTables("create table IF not exists qepoxy2.PriceData(PrName varchar(100), Price double, Weight double, Volume double,Coverage double,Unit char(15),prDate char(15), primary key(PrName), unique(PrName));");
            createTables("create table IF not exists qepoxy2.fslink(FsName varchar(100), PaName varchar(100))");
            createTables("create table IF not exists qepoxy2.packlink(PaName varchar(100), PrName varchar(100))");            
        }

        private void createTables(string sqlQueryin)
        {
            try
            {
                sqlConn.Open();
                sqlQuery = sqlQueryin;
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void attemptConnect()  
        {
            try
            {
                sqlConn.Open();
                serverResult = true;
                sqlConn.Close();

            }
            catch(Exception ex)
            {
                string msg = ex.Message;
                serverResult = false;                
            }
        }

        public void ImportAllDataFromDiskToSQL(string filepath)
        {
            SystemsList = LoadFSData(filepath);
            PackList = LoadPackData(filepath);
            PriceList = LoadProductData(filepath);
            LoadAdditions(filepath);
            LoadHrRates(filepath);

            //---------------------------------------------------------------------------------------------
            // loading the Hrratessettings into SQL data base
            //---------------------------------------------------------------------------------------------            
            createTables("create table IF not exists qepoxy2.hrratesettings(HrName varchar(100), HrWeight char(20), primary key(HrName), unique(HrName));");
            SQLErrorList.Add("---------------- Hrratessettings table Finished ----------------");

            //---------------------------------------------------------------------------------------------
            // loading the Additions into SQL data base
            //---------------------------------------------------------------------------------------------            
            createTables("create table IF not exists qepoxy2.additions(ItemName varchar(50));");
            SQLErrorList.Add("---------------- Additions table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            // loading the fsdata into SQL data base
            //---------------------------------------------------------------------------------------------           

            createTables("create table IF not exists qepoxy2.FSdata(FsName varchar(100), Type char(20), primary key(FsName), unique(FsName))");
            SQLErrorList.Add("---------------- FSdata table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            // loading the Packdata into SQL data base
            //---------------------------------------------------------------------------------------------
            createTables("create table IF not exists qepoxy2.PackData(PaName varchar(100), Price double, Weight double, Volume double,Coverage double,Unit char(15),primary key(PaName), unique(PaName));");
            SQLErrorList.Add("---------------- PackData table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            // loading the PriceData into SQL data base
            //---------------------------------------------------------------------------------------------
            createTables("create table IF not exists qepoxy2.PriceData(PrName varchar(100), Price double, Weight double, Volume double,Coverage double,Unit char(15),prDate char(15), primary key(PrName), unique(PrName));");
            SQLErrorList.Add("---------------- PriceData table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            //  loading the FSLink into SQL data base
            //  Sets up the links between the System names and the Pack names using foreign keys
            //---------------------------------------------------------------------------------------------
            createTables("create table IF not exists qepoxy2.fslink(FsName varchar(100), PaName varchar(100))");
            SQLErrorList.Add("---------------- FSLink table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            //  loading the PackLink into SQL data base
            //  Sets up the links between the Pack names and the product names using foreign keys
            //---------------------------------------------------------------------------------------------
            createTables("create table IF not exists qepoxy2.packlink(PaName varchar(100), PrName varchar(100))");
            SQLErrorList.Add("---------------- PackLink table Finished ----------------");
            //---------------------------------------------------------------------------------------------
            //      Population of Lists into DataSheets setup above 
            //---------------------------------------------------------------------------------------------
            // Loop through systemlist to add data to fsdata
            //---------------------------------------------------------------------------------------------
            SQLErrorList.Add("---------------- Start Population of FSdata table ----------------");
            foreach (ProductSystem psys in SystemsList)
            {
                try
                {
                    sqlConn.Open();                    
                    sqlQuery = "insert into qepoxy2.fsdata values('" + psys.BaseMat + "','" + psys.Type + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                foreach (Pack pa in psys.PackList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.fslink values('" + psys.BaseMat + "','" + pa.Name + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("--> " + msg);
                        sqlConn.Close();
                    }
                }
            }

            //---------------------------------------------------------------------------------------------
            // loop through Packdata to add items to sql Db
            //---------------------------------------------------------------------------------------------
            SQLErrorList.Add("---------------- Start Population of PackData table ----------------");
            foreach (Pack paItem in PackList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.packData values('" + paItem.Name + "','" + paItem.Price + "','" + paItem.Weight + "','" + paItem.Volume + "','" + paItem.Coverage + "','" + paItem.Unit + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }                
                foreach (Priceitem pi in paItem.PriceItemList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.packlink values('" + paItem.Name + "','" + pi.Name + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }
                }
            }

            //---------------------------------------------------------------------------------------------
            // loop through Product data to add items to sql Db
            //---------------------------------------------------------------------------------------------
            SQLErrorList.Add("---------------- Start Population of PriceData table ----------------");
            foreach (Priceitem prItem in PriceList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.PriceData values('" + prItem.Name + "','" + prItem.Price + "','" + prItem.Weight + "','" + prItem.Volume + "','" + prItem.Coverage + "','" + prItem.Unit + "','" + prItem.productDate + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }

            //---------------------------------------------------------------------------------------------
            // loop through Product data to add items to sql Db
            //---------------------------------------------------------------------------------------------
            SQLErrorList.Add("---------------- Start Population of HrratesSettings table ----------------");
            foreach (Hrrates HrItem in HrrateList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.hrratesettings values('" + HrItem.hrsname + "','" + HrItem.hrsWeight + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }

            SQLErrorList.Add("%%% ---------------- Done ---------------- %%%");
        

        //---------------------------------------------------------------------------------------------
        // loop through Additions to add items to sql Db
        //---------------------------------------------------------------------------------------------
        SQLErrorList.Add("---------------- Start Population of Additions table ----------------");
            foreach (string strItem in AdditionsList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.additions values('" + strItem +  "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();

                    }
                    catch (Exception ex)
                    {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }

            SQLErrorList.Add("%%% ---------------- Done ---------------- %%%");
        }

        
        //-------------------------------------------------------------------
        //---- attempt to rebuild the lists using mySQL calls to remote Db --- 
        //---- FS data population -----
        //-------------------------------------------------------------------
        public void RebuildProductSystems()
        {
            //SystemsList.Clear();
            //PackList.Clear();
            //PriceList.Clear();            
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.fsdata";
                sqlRd = sqlCmd.ExecuteReader();
                SystemsList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.
                        ProductSystem psItem = new ProductSystem();
                        psItem.BaseMat = sqlRd.GetString(sqlRd.GetOrdinal("FsName"));
                        psItem.Type = sqlRd.GetString(sqlRd.GetOrdinal("Type"));
                        SystemsList.Add(psItem);
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
            RelinkSystemDataLinks();
        }
        //-------------------------------------------------------------------
        //---- Packlist data population--- 
        //-------------------------------------------------------------------
        public void RebuildPackData()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.packdata";
                sqlRd = sqlCmd.ExecuteReader();
                PackList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.                        
                        Pack pAItem = new Pack();
                        pAItem.Name = sqlRd.GetString(sqlRd.GetOrdinal("PaName"));
                        pAItem.Price = sqlRd.GetDouble(sqlRd.GetOrdinal("Price"));
                        pAItem.Weight = sqlRd.GetDouble(sqlRd.GetOrdinal("Weight"));
                        pAItem.Volume = sqlRd.GetDouble(sqlRd.GetOrdinal("Volume"));
                        pAItem.Coverage = sqlRd.GetDouble(sqlRd.GetOrdinal("Coverage"));
                        pAItem.Unit = sqlRd.GetString(sqlRd.GetOrdinal("Unit"));                        
                        PackList.Add(pAItem);
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.packlink";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.                        
                        string pkName = sqlRd.GetString(sqlRd.GetOrdinal("PaName"));
                        string PrName = sqlRd.GetString(sqlRd.GetOrdinal("PrName"));

                        foreach (Pack pk in PackList)
                        {
                            if (pk.Name == pkName)
                            {
                                foreach (Priceitem prItem in PriceList)
                                {
                                    if (prItem.Name == PrName)
                                    {
                                        Priceitem pitem = new Priceitem();
                                        pitem = prItem;
                                        pk.PriceItemList.Add(pitem);
                                    }
                                }
                            }
                        }
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }


        }               
        //-------------------------------------------------------------------
        //---- Pricelist data population--- 
        //-------------------------------------------------------------------
        public void RebuildPriceData()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.pricedata";
                sqlRd = sqlCmd.ExecuteReader();
                PriceList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.
                        Priceitem prItem = new Priceitem();
                        prItem.Name = sqlRd.GetString(sqlRd.GetOrdinal("PrName"));
                        prItem.Price = sqlRd.GetDouble(sqlRd.GetOrdinal("Price"));
                        prItem.Weight = sqlRd.GetDouble(sqlRd.GetOrdinal("Weight"));
                        prItem.Volume = sqlRd.GetDouble(sqlRd.GetOrdinal("Volume"));
                        prItem.Coverage = sqlRd.GetDouble(sqlRd.GetOrdinal("Coverage"));
                        prItem.Unit = sqlRd.GetString(sqlRd.GetOrdinal("Unit"));
                        prItem.productDate = sqlRd.GetString(sqlRd.GetOrdinal("prDate"));
                        PriceList.Add(prItem);
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }

        // ----------------------------------------------------------------
        // attempt relinking into lists
        // ----------------------------------------------------------------
        public void RelinkPackLinks()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.packlink";
                sqlRd = sqlCmd.ExecuteReader();
                foreach (Pack pk in PackList)
                {
                    pk.PriceItemList.Clear();
                }

                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.                        
                        string pkName = sqlRd.GetString(sqlRd.GetOrdinal("PaName"));
                        string PrName = sqlRd.GetString(sqlRd.GetOrdinal("PrName"));
                        
                        foreach (Pack pk in PackList)
                        {
                            if (pk.Name == pkName)
                            {                                
                                foreach (Priceitem prItem in PriceList)
                                {
                                    if (prItem.Name == PrName)
                                    {
                                        Priceitem pitem = new Priceitem();
                                        pitem = prItem;
                                        pk.PriceItemList.Add(pitem);
                                    }
                                }
                            }
                        }
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void RelinkSystemDataLinks()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.fslink";
                sqlRd = sqlCmd.ExecuteReader();
                foreach (ProductSystem psys in SystemsList)
                {
                    psys.PackList.Clear();
                }

                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.                        
                        string sysName = sqlRd.GetString(sqlRd.GetOrdinal("FsName"));
                        string pkName = sqlRd.GetString(sqlRd.GetOrdinal("PaName"));
                        foreach (ProductSystem psys in SystemsList)
                        {
                            if (psys.BaseMat == sysName)
                            {//pk.PriceItemList.Clear();
                                foreach (Pack np in PackList)
                                {
                                    if (np.Name == pkName)
                                    {
                                        psys.PackList.Add(np);
                                    }
                                }
                            }
                        }
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        // ----------------------------------------------------------------
        // Reload All SQL lists and link them
        // ----------------------------------------------------------------
        public void LoadSQLData()
        {           
            RebuildProductSystems();
            RebuildPackData();
            RebuildPriceData();
            RebuildQuoteList();
            RebuildCompanyList();
            RebuildQuoteDBLists();
            RebuildGenQuoteDBLists();
            RelinkPackLinks();
            RelinkSystemDataLinks();
            RebuildHrRates();
            RebuildAdditions();
        }

        public void sqlDropTable(string text)
        {
            try
            {
                sqlConn.Open();
                //sqlQuery = "DROP TABLE IF EXISTS qepoxy2.commentsdb;";
                sqlCmd = new MySqlCommand(text, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }
        }

        public void DropAllTables()
        {

            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.commentsdb;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.hrratesettings;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.additions;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.companydata;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.fslink;"); 
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.packlink;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.fsdata;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.packdata;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.pricedata;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.quotecomms;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.quotedata;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.quote_blockfs;");
            sqlDropTable("DROP TABLE IF EXISTS qepoxy2.quoteoptions;");
           
        }
        public void AdditemToSystemsList(ProductSystem Pimport)
        {
            bool Inlist = false;
            foreach (ProductSystem psItem in SystemsList)
            {
                if (Pimport.BaseMat == psItem.BaseMat)
                {
                    Inlist = true;
                    //sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                    //
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.fsdata set FsName = @Name, Type = @Type where FsName = @Name";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Pimport.BaseMat);
                        sqlCmd.Parameters.AddWithValue("@Type", Pimport.Type);

                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "delete from qepoxy2.fslink where FsName = @Name";
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Pimport.BaseMat);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }

                    foreach (Pack pai in Pimport.PackList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.fslink values('" + Pimport.BaseMat + "','" + pai.Name + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            //sqlCmd.ExecuteReader();
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add("-> " + msg);
                            sqlConn.Close();
                        }
                    }
                }
            }
            if (!Inlist)
            {
                // sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.fsdata (FsName, Type) values('" + Pimport.BaseMat + "','" + Pimport.Type + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                foreach (Pack pai in Pimport.PackList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.fslink values('" + Pimport.BaseMat + "','" + pai.Name + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }
                }
            }
        }
        public void AdditemToPackList(Pack Pimport)
        {
            bool Inlist = false;
            foreach (Pack paItem in PackList)
            {
                if (Pimport.Name == paItem.Name)
                {
                    Inlist = true;
                    //sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                    
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.packdata set PaName = @Name, Price = @Price," +
                            "Weight = @Weight,Volume = @Volume,Coverage = @Coverage,Unit = @Unit " +
                            "where PaName = @Name";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Pimport.Name);
                        sqlCmd.Parameters.AddWithValue("@Price", Pimport.Price);
                        sqlCmd.Parameters.AddWithValue("@Weight", Pimport.Weight);
                        sqlCmd.Parameters.AddWithValue("@Volume", Pimport.Volume);
                        sqlCmd.Parameters.AddWithValue("@Coverage", Pimport.Coverage);
                        sqlCmd.Parameters.AddWithValue("@Unit", Pimport.Unit);

                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "delete from qepoxy2.packlink where PaName = @Name";
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Pimport.Name);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }
                    foreach (Priceitem pi in Pimport.PriceItemList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.packlink values('" + Pimport.Name + "','" + pi.Name + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            //sqlCmd.ExecuteReader();
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add("-> " + msg);
                            sqlConn.Close();
                        }
                    }
                }
            }
            if (!Inlist)
            {
                // sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                try
                {
                    sqlConn.Open();
                    MySqlCommand sqlCmd = new MySqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = "delete from qepoxy2.packlink where PaName = @Name";
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@Name", Pimport.Name);
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }

                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.packdata (PaName, Price, Weight, Volume, Coverage, Unit) values('" + Pimport.Name + "','" + Pimport.Price + "','" + Pimport.Weight + "','" + Pimport.Volume + "','" + Pimport.Coverage + "','" + Pimport.Unit + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                foreach (Priceitem pi in Pimport.PriceItemList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.packlink values('" + Pimport.Name + "','" + pi.Name + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }
                }
            }
        }
        public void AdditemToPriceList(Priceitem Pimport)
        {
            bool Inlist = false;           
            foreach (Priceitem prItem in PriceList)
            {
                if (Pimport.Name == prItem.Name)
                {
                    Inlist = true;
                    //sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.pricedata set PrName = @Name, Price = @Price," +
                            "Weight = @Weight,Volume = @Volume,Coverage = @Coverage,Unit = @Unit, prDate = @prodDate " +
                            "where PrName = @Name";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Pimport.Name);
                        sqlCmd.Parameters.AddWithValue("@Price", Pimport.Price);
                        sqlCmd.Parameters.AddWithValue("@Weight", Pimport.Weight);
                        sqlCmd.Parameters.AddWithValue("@Volume", Pimport.Volume);
                        sqlCmd.Parameters.AddWithValue("@Coverage", Pimport.Coverage);
                        sqlCmd.Parameters.AddWithValue("@Unit", Pimport.Unit);
                        sqlCmd.Parameters.AddWithValue("@prodDate", Pimport.productDate);

                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                    
                }
            }
            if (!Inlist)
            {                
                try
                {
                    sqlConn.Open();
                    MySqlCommand sqlCmd = new MySqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = "delete from qepoxy2.packlink where PaName = @Name";
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@Name", Pimport.Name);
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.pricedata (PrName, Price, Weight, Volume, Coverage, Unit,prDate) values('" + Pimport.Name + "','" + Pimport.Price + "','" + Pimport.Weight + "','" + Pimport.Volume + "','" + Pimport.Coverage + "','" + Pimport.Unit + "','" + Pimport.productDate + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }                
            }
            RebuildPackData();
        }
        public void AddItemToCompanyList(Company Cin)
        {
            bool Inlist = false;
            foreach (Company cItem in Complist)
            {
                if (Cin.CompName == cItem.CompName)
                {
                    Inlist = true;
                    //sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.companydata set CName = @Name, EmailAddr = @Email," +
                            "Address = @Addr, Suburb = @suburb, State = @state, Postcode = @pc, ContactName = @contactname,ContactNum = @contactnum" +
                            "where CName = @Name";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Cin.CompName);
                        sqlCmd.Parameters.AddWithValue("@Email", Cin.Email);
                        sqlCmd.Parameters.AddWithValue("@suburb", Cin.Suburb);
                        sqlCmd.Parameters.AddWithValue("@state", Cin.State);
                        sqlCmd.Parameters.AddWithValue("@pc", Cin.PostCode);
                        sqlCmd.Parameters.AddWithValue("@contactname", Cin.ContactName);
                        sqlCmd.Parameters.AddWithValue("@contactnum", Cin.ContactNum);                        

                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                }
            }
            if (!Inlist)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.companydata (CName, EmailAddr, Address, Suburb, State, Postcode,ContactName,ContactNum) values('" +
                        Cin.CompName + "','" + Cin.Email + "','" + Cin.Address + "','" + Cin.Suburb + "','" + Cin.State + "','" + Cin.PostCode + "','" + Cin.ContactName + "','" + Cin.ContactNum + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }
        }
        public void DelItemFromPriceList(string Pimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.pricedata where PrName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Pimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void DelItemFromPackList(string Paimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.packlink where PaName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Paimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.packdata where PaName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Paimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void DelItemFromSystemsList(string Fsimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.fslink where FsName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Fsimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.fsdata where FsName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Fsimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

        }
        public void RebuildQuoteDBLists()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.Commentsdb";
                sqlRd = sqlCmd.ExecuteReader();
                QuoteDbList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        string Q1Headingstr = sqlRd.GetString(sqlRd.GetOrdinal("Heading"));
                        string Q1commentstr = sqlRd.GetString(sqlRd.GetOrdinal("Comment"));

                        bool inList = false;
                        foreach (QuoteComment Qcom in QuoteDbList)
                        {
                            if (Qcom.QcHeading == Q1Headingstr)
                            {
                                inList = true;
                                Qcom.QcTextList.Add(Q1commentstr);
                            }
                        }
                        if (!inList)
                        {
                            QuoteComment Qcom = new QuoteComment();
                            Qcom.QcHeading = Q1Headingstr;
                            Qcom.QcTextList.Add(Q1commentstr);
                            QuoteDbList.Add(Qcom);
                        }
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void RebuildGenQuoteDBLists()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.gencomments";
                sqlRd = sqlCmd.ExecuteReader();
                GCommentdbList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {                    
                        string CommentStr = sqlRd.GetString(sqlRd.GetOrdinal("Comment"));
                        GCommentdbList.Add(CommentStr);
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void ImportQuotesFromDiskToSQL(string filepath)
        {
            try
            {
                sqlConn.Open();
                sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quoteoptions;";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quotecomms;";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quote_blockfs;";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }
            //try
            //{
            //    sqlConn.Open();
            //    sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quotedata;";
            //    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
            //    sqlCmd.ExecuteReader();
            //    sqlCmd.ExecuteNonQuery();
            //    sqlConn.Close();
            //}
            //catch (Exception ex)
            //{
            //    string msg = ex.Message;
            //    SQLErrorList.Add("-> " + msg);
            //    sqlConn.Close();
            //}
            try
            {
                sqlConn.Open();
                sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quote_blockdata;";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                sqlQuery = "DROP TABLE IF EXISTS qepoxy2.quote_hrsdata;";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add("-> " + msg);
                sqlConn.Close();
            }
            List<Quote_Block> QList = LoadQuoteFSdata(filepath);

            //---------------------------------------------------------------------------------------------
            // loading the Quote Comment data into SQL data base
            //---------------------------------------------------------------------------------------------            
            try
            {
                SQLErrorList.Add("---------------- create Quote Form Options table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.quoteoptions(QName varchar(100), jobnum char(30), jobDate char(30), companyName char(30),email char(30),Address char(100), " +
                    "Suburb char(30),State char(30), Postcode char(30), clientName char(30),clientNum char(30),project char(100),totalEst char(30)," +
                                      "Q1Allowance char(30),Q1Lm char(30),Q1Days char(30),Q1Heading char(100),Q1SlipRate char(30),Q1StageVal char(30),Q1Warrenty char(30)," +
                    "Q2Enable varchar(20), Q2Allowance char(30), Q2Lm char(30), Q2Days char(30), Q2Heading char(100), Q2SlipRate char(30), Q2StageVal char(30), Q2Warrenty char(30)," +
                    "Q3Enable varchar(20), Q3Allowance char(30), Q3Lm char(30), Q3Days char(30), Q3Heading char(100), Q3SlipRate char(30), Q3StageVal char(30), Q3Warrenty char(30)," +
                    "Rubbish varchar(20), Phase3 varchar(20), Drytime varchar(20), JointColour varchar(20), Colour varchar(20), SlipRes varchar(20), QuoteStages varchar(20), ArrowHead varchar(20), AllowCtrlJoins varchar(20)," +
                    "jointCstr varchar(30), Colourstr varchar(30), slipR varchar(30), QuoteStagesStr varchar(30), ArrowH varchar(30), AllowCtrl varchar(30));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            SQLErrorList.Add("---------------- Quote Form Options Lists table Finished ----------------");
            try
            {
                SQLErrorList.Add("---------------- create FSdata table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.quotecomms(QName varchar(100),Listtype varchar(30), Heading varchar(100), subHeading char(20), Comment varchar(200));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            SQLErrorList.Add("---------------- Quote Comments table Finished ----------------");

            //---------------------------------------------------------------------------------------------
            // loading the Quote into SQL data base
            //---------------------------------------------------------------------------------------------            
            try
            {
                SQLErrorList.Add("---------------- Create quotefs Table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.quote_blockfs(Qname varchar(100), HrRate double, HrAddition double, Additions double,Markup double, Date varchar(20), primary key(Qname), unique(Qname))";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            //---------------------------------------------------------------------------------------------
            // build the QuoteLists into SQL data base table
            //---------------------------------------------------------------------------------------------    
            try
            {
                SQLErrorList.Add("---------------- Create quoteData Table----------------");
                sqlConn.Open();
                //"Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit
                sqlQuery = "create table IF not exists qepoxy2.quotedata(Qname varchar(100), Type char(20) ,PackId int, PackName varchar(100), PiName varchar(100), Price double, Weight double, Volume double, Meterage double,Coverage double,unit char (10));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            
            try
            {
                SQLErrorList.Add("---------------- Create quote_blockdata Table----------------");
                sqlConn.Open();
                //"Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit
                sqlQuery = "create table IF not exists qepoxy2.quote_blockdata(Qname varchar(30), Type varchar(30) ,PackId int, InputName varchar(100), Ammount varchar(100), Area varchar(30), Lm varchar(30), Cost varchar(30),PackAmt varchar(30)," +
                    "PackSystemName varchar(20),Packvol varchar(30), Price varchar(30), Stage varchar(30), Symbol varchar(30), Systems varchar(30));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                SQLErrorList.Add("---------------- Create quote_hrsdata Table----------------");
                sqlConn.Open();
                //"Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit
                sqlQuery = "create table IF not exists qepoxy2.quote_hrsdata(Qname varchar(30), Type varchar(30) ,PackId int, Hours varchar(10), Lm varchar(100), Area varchar(30), Method varchar(30), Stage varchar(30),Systems varchar(30),Techhrs varchar(20));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close(); 
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            //---------------------------------------------------------------------------------------------
            //      Population of Quotes into DataSheets setup above 
            //---------------------------------------------------------------------------------------------
            //      Loop through QuoteFs to add data to quotefs table
            //---------------------------------------------------------------------------------------------
            SQLErrorList.Add("---------------- Start Population of quotefs table ----------------");
            
            foreach (Quote_Block Qbl in QList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.quote_blockfs values('" + Qbl.QName + "','" + Qbl.savedHrRate + "','" + Qbl.savedHradd + "','" + Qbl.savedAdds + "','" + Qbl.savedMarkup + "','" + Qbl.SaveDate + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                AddtoQuote_BlockList(Qbl);


                DataloadQuoteData(Qbl,filepath);
                foreach (string qinfostr in Qbl.Qinfo.CommentList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','general','','','" + qinfostr + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                }
                foreach (QuoteComment Qcomm in Qbl.Qinfo.Q1CommentList)
                {
                    foreach (string str in Qcomm.QcTextList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q1','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            //sqlCmd.ExecuteReader();
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add(msg);
                            sqlConn.Close();
                        }
                    }
                }
                foreach (QuoteComment Qcomm in Qbl.Qinfo.Q2CommentList)
                {
                    foreach (string str in Qcomm.QcTextList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q2','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            //sqlCmd.ExecuteReader();
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add(msg);
                            sqlConn.Close();
                        }
                    }
                }
                foreach (QuoteComment Qcomm in Qbl.Qinfo.Q3CommentList)
                {
                    foreach (string str in Qcomm.QcTextList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q3','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            //sqlCmd.ExecuteReader();
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add(msg);
                            sqlConn.Close();
                        }
                    }
                }
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.quoteoptions values( '" + Qbl.QName + "','" + Qbl.Qinfo.Jobnum + "','" + Qbl.Qinfo.JobDate + "','" + Qbl.Qinfo.CompanyName + "','" + Qbl.Qinfo.Email + "','" + Qbl.Qinfo.Address + "','" + Qbl.Qinfo.Suburb + "','" + Qbl.Qinfo.State + "','" + Qbl.Qinfo.Postcode + "','" + Qbl.Qinfo.ClientName + "','" + Qbl.Qinfo.ClientNum + "','" + Qbl.Qinfo.Project + "','" + Qbl.Qinfo.TotalEst + "','" +
                                                                        Qbl.Qinfo.Q1Allowance + "','" + Qbl.Qinfo.Q1Lm + "','" + Qbl.Qinfo.Q1Days + "','" + Qbl.Qinfo.Q1Heading + "','" + Qbl.Qinfo.Q1SlipRate + "','" + Qbl.Qinfo.Q1StageVal + "','" + Qbl.Qinfo.Q1Warrenty + "','" +
                                           Qbl.Qinfo.Q2Enable + "','" + Qbl.Qinfo.Q2Allowance + "','" + Qbl.Qinfo.Q2Lm + "','" + Qbl.Qinfo.Q2Days + "','" + Qbl.Qinfo.Q2Heading + "','" + Qbl.Qinfo.Q2SlipRate + "','" + Qbl.Qinfo.Q2StageVal + "','" + Qbl.Qinfo.Q2Warrenty + "','" +
                                           Qbl.Qinfo.Q3Enable + "','" + Qbl.Qinfo.Q3Allowance + "','" + Qbl.Qinfo.Q3Lm + "','" + Qbl.Qinfo.Q3Days + "','" + Qbl.Qinfo.Q3Heading + "','" + Qbl.Qinfo.Q3SlipRate + "','" + Qbl.Qinfo.Q3StageVal + "','" + Qbl.Qinfo.Q3Warrenty + "','" +
                                           Qbl.Qinfo.Rubbish + "','" + Qbl.Qinfo.Phase3 + "','" + Qbl.Qinfo.Drytime + "','" + Qbl.Qinfo.JointColour + "','" + Qbl.Qinfo.Colour + "','" + Qbl.Qinfo.SlipRes + "','" + Qbl.Qinfo.QuoteStages + "','" + Qbl.Qinfo.ArrowHead + "','" + Qbl.Qinfo.AllowCtrlJoins + "','" +
                                           Qbl.Qinfo.JointCstr + "','" + Qbl.Qinfo.Colstr + "','" + Qbl.Qinfo.SlipRstr + "','" + Qbl.Qinfo.QuoteStagesstr + "','" + Qbl.Qinfo.Arrow + "','" + Qbl.Qinfo.AllowCtrlstr + "');";

                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                //foreach (Pack pitem in Qbl.QFPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Floor','"+ ID +"','" + pitem.Name + "','none','" + pitem.Price + "','"
                //                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";                       
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Floor','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }
                //}
                //foreach (Pack pitem in Qbl.QCPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Coving','" + ID + "','" + pitem.Name + "','none','" + pitem.Price + "','" + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Coving','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }

                //}
                //foreach (Pack pitem in Qbl.QSFPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Subfill','" + ID + "','" + pitem.Name + "','none','" + pitem.Price + "','" + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Subfill','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }
                //}
                //foreach (Pack pitem in Qbl.QLMPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','LineMarking','" + ID + "','" + pitem.Name + "','none','" + pitem.Price + "','" + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','LineMarking','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }
                //}
                //foreach (Pack pitem in Qbl.QCVPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Chevron','" + ID + "','" + pitem.Name + "','none','" + pitem.Price + "','" + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','Chevron','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }
                //}
                //foreach (Pack pitem in Qbl.QPCPackList)
                //{
                //    try
                //    {
                //        sqlConn.Open();
                //        ID++;
                //        sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','PolishedCement','" + ID + "','" + pitem.Name + "','none','" + pitem.Price + "','" + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //        //sqlCmd.ExecuteReader();
                //        sqlCmd.ExecuteNonQuery();
                //        sqlConn.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        string msg = ex.Message;
                //        SQLErrorList.Add(msg);
                //        sqlConn.Close();
                //    }
                //    foreach (Priceitem prItem in pitem.PriceItemList)
                //    {
                //        try
                //        {
                //            sqlConn.Open();
                //            sqlQuery = "insert into qepoxy2.quotedata values('" + Qbl.QName + "','PolishedCement','" + ID + "','" + pitem.Name + "','" + prItem.Name + "','" + +pitem.Price + "','"
                //                                        + pitem.Weight + "','" + pitem.Volume + "','" + pitem.Meterage + "','" + pitem.Coverage + "','" + pitem.Unit + "');";
                //            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //            //sqlCmd.ExecuteReader();
                //            sqlCmd.ExecuteNonQuery();
                //            sqlConn.Close();
                //        }
                //        catch (Exception ex)
                //        {
                //            string msg = ex.Message;
                //            SQLErrorList.Add(msg);
                //            sqlConn.Close();
                //        }
                //    }
                //}
            }
        }
        public void AddtoGenComments(string instr)
        {
            try
            {
                sqlConn.Open();
                sqlQuery = "insert into qepoxy2.gencomments values('" + instr + "');";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void AddtoCommentsDb(QuoteComment inQC)
        {
            bool inList = false;
            foreach(QuoteComment qcl in QuoteDbList)
            {
                if(inQC.QcHeading == qcl.QcHeading)
                {
                    inList = true;
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "delete from qepoxy2.commentsdb where Heading = @Head";
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Head", inQC.QcHeading);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                    foreach (string str in inQC.QcTextList)
                    {
                        try
                        {
                            sqlConn.Open();
                            sqlQuery = "insert into qepoxy2.commentsdb values('" + inQC.QcHeading + "','" + inQC.QcSubHeading + "','" + str + "');";
                            sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                            sqlCmd.ExecuteNonQuery();
                            sqlConn.Close();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            SQLErrorList.Add(msg);
                            sqlConn.Close();
                        }
                    }


                }
            }
            if(!inList)
            {
                foreach (string str in inQC.QcTextList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.CommentsDB values('" + inQC.QcHeading + "','" + inQC.QcSubHeading + "','" + str + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                }
            }  
        }

        //---------------------------------------------------------------------------------------------------------------------
        // Comments DB list
        //---------------------------------------------------------------------------------------------------------------------
        public void ImportCommentDbToSql(string filepath)
        {
            //public List<QuoteComment> QuoteDbList = new List<QuoteComment>();
            //public List<String> GCommentdbList = new List<String>();
            QuoteDbList.Clear();
            GCommentdbList.Clear();

            LoadDbFromDisk(filepath);
            LoadGCommentsDb(filepath);
            try
            {
                SQLErrorList.Add("---------------- create General comments table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.Gencomments(Comment varchar(200));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            //< QuoteComment > QuoteDbList

            SQLErrorList.Add("---------------- General comments table Finished ----------------");
            try
            {
                SQLErrorList.Add("---------------- create CommentsDB table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.CommentsDB( Heading varchar(200), subHeading char(20), Comment varchar(300));";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }


            SQLErrorList.Add("----------------  CommentsDB table Finished ----------------");
            foreach (string str in GCommentdbList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.Gencomments values('" + str + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);                    
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }

            foreach (QuoteComment qc in QuoteDbList)
            {
                foreach (string str in qc.QcTextList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.CommentsDB values('" + qc.QcHeading + "','" + qc.QcSubHeading + "','" + str + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                }
            }
        }
        public void ImportCommentListsToSql()
        {
            //LoadCompanyData();

            foreach (Company Citem in Complist)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.companydata values('" + Citem.CompName + "','" + Citem.Email + "','" + Citem.Address + "','" + Citem.Suburb + "','" + Citem.State + "','" + Citem.PostCode + "','" + Citem.ContactName + "','" + Citem.ContactNum + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }
            RebuildCompanyList();
        }
        public void ImportCompanyDataToSql(string filepath)
        {
            try
            {
                SQLErrorList.Add("---------------- create Company table----------------");
                sqlConn.Open();
                sqlQuery = "create table IF not exists qepoxy2.companydata(CName varchar(100), EmailAddr varchar(50), Address varchar(50), Suburb varchar(50), State varchar(10), Postcode varchar(10), ContactName varchar(50), ContactNum varchar(50)); ";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlCmd.ExecuteReader();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            LoadCompanyData(filepath);

            foreach(Company Citem in Complist)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.companydata values('" + Citem.CompName + "','" + Citem.Email + "','" + Citem.Address + "','" + Citem.Suburb + "','" + Citem.State + "','" + Citem.PostCode + "','" + Citem.ContactName + "','" + Citem.ContactNum + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }
            RebuildCompanyList();
        }
        public void DelItemFromCompanyList(string Cimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.companydata where CName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Cimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void DelItemFromCommentDb(string strhead)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.commentsdb where Heading = @Head";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Head", strhead);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void DelItemFromGenComms(string Instr)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.gencomments where Comment = @Comm";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Comm", Instr);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void RebuildCompanyList()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.companydata";
                sqlRd = sqlCmd.ExecuteReader();
                Complist.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        Company compIn = new Company();
                        compIn.CompName = sqlRd.GetString(sqlRd.GetOrdinal("CName"));
                        compIn.Email = sqlRd.GetString(sqlRd.GetOrdinal("EmailAddr"));
                        compIn.Address = sqlRd.GetString(sqlRd.GetOrdinal("Address"));
                        compIn.Suburb = sqlRd.GetString(sqlRd.GetOrdinal("Suburb"));
                        compIn.State = sqlRd.GetString(sqlRd.GetOrdinal("State"));
                        compIn.PostCode = sqlRd.GetString(sqlRd.GetOrdinal("Postcode"));
                        compIn.ContactName = sqlRd.GetString(sqlRd.GetOrdinal("ContactName"));
                        compIn.ContactNum = sqlRd.GetString(sqlRd.GetOrdinal("ContactNum"));
                        Complist.Add(compIn);
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void RebuildQuoteList()
        {            
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_blockfs";
                sqlRd = sqlCmd.ExecuteReader();
                Q_bList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //quotefs(Qname varchar(100), HrRate double, HrAddition double, Additions double, Date varchar(20), primary key(Qname), unique(Qname))";
                        //I would also check for DB.Null here before reading the value.
                        Quote_Block qbitem = new Quote_Block();
                        qbitem.QName = sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                        qbitem.savedHrRate = sqlRd.GetString(sqlRd.GetOrdinal("HrRate"));
                        qbitem.savedHradd = sqlRd.GetString(sqlRd.GetOrdinal("HrAddition"));
                        qbitem.savedAdds = sqlRd.GetString(sqlRd.GetOrdinal("additions"));
                        qbitem.SaveDate = sqlRd.GetString(sqlRd.GetOrdinal("Date"));
                        qbitem.savedMarkup = sqlRd.GetString(sqlRd.GetOrdinal("Markup"));
                        Q_bList.Add(qbitem);
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_blockdata";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.
                        string QName = sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                        string Type = sqlRd.GetString(sqlRd.GetOrdinal("Type"));
                        int Packidentifier = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                        string qPackname = sqlRd.GetString(sqlRd.GetOrdinal("InputName"));
                        double qAmmount = sqlRd.GetDouble(sqlRd.GetOrdinal("Ammount"));
                        double qArea = sqlRd.GetDouble(sqlRd.GetOrdinal("Area"));
                        double qLm = sqlRd.GetDouble(sqlRd.GetOrdinal("Lm"));
                        double qCost = sqlRd.GetDouble(sqlRd.GetOrdinal("Cost"));
                        double qPackAmt = sqlRd.GetDouble(sqlRd.GetOrdinal("PackAmt"));
                        string qPackSystemName = sqlRd.GetString(sqlRd.GetOrdinal("PackSystemName"));
                        double qPackvol = sqlRd.GetDouble(sqlRd.GetOrdinal("Packvol"));
                        string qStage = sqlRd.GetString(sqlRd.GetOrdinal("Stage"));
                        string qSymbol = sqlRd.GetString(sqlRd.GetOrdinal("Symbol"));
                        string qSystems = sqlRd.GetString(sqlRd.GetOrdinal("Systems"));
                        double qPrice = sqlRd.GetDouble(sqlRd.GetOrdinal("Price"));

                        foreach (Quote_Block Qbl in Q_bList)
                        {
                            if (Qbl.QName == QName)
                            {
                                if (Type == "stage")
                                {
                                    CSstage CSitem = new CSstage();
                                    CSitem.inputname = qPackname;
                                    CSitem.Ammount = qAmmount;
                                    CSitem.Area = qArea;
                                    CSitem.Lm = qLm;
                                    CSitem.cost = qCost;
                                    CSitem.packsAmnt = qPackAmt;
                                    CSitem.PackSystemName = qPackSystemName;
                                    CSitem.Packvol = qPackvol;
                                    CSitem.Staging = qStage;
                                    CSitem.symbol = qSymbol;
                                    CSitem.Systems = qSystems;
                                    CSitem.price = qPrice;
                                    Qbl.qbStagelist.Add(CSitem);
                                }                                
                            }                           
                        }                        
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_hrsdata";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.
                        string QName = sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                        string Type = sqlRd.GetString(sqlRd.GetOrdinal("Type"));
                        int Packidentifier = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                        string qhours = sqlRd.GetString(sqlRd.GetOrdinal("Hours"));
                        string qlm = sqlRd.GetString(sqlRd.GetOrdinal("Lm"));
                        string qArea = sqlRd.GetString(sqlRd.GetOrdinal("Area"));
                        string qmethod = sqlRd.GetString(sqlRd.GetOrdinal("Method"));
                        string qStagename = sqlRd.GetString(sqlRd.GetOrdinal("Stage"));
                        string qsystemName = sqlRd.GetString(sqlRd.GetOrdinal("Systems"));
                        string qtechhrs = sqlRd.GetString(sqlRd.GetOrdinal("Techhrs"));

                        foreach (Quote_Block Qbl in Q_bList)
                        {
                            if (Qbl.QName == QName)
                            {                               
                                EstHoursSave HRitem = new EstHoursSave();
                                HRitem.Hours = qhours;
                                HRitem.lm = qlm;
                                HRitem.m2 = qArea;
                                HRitem.method = qmethod;
                                HRitem.stageName = qStagename;
                                HRitem.systemName = qsystemName;
                                HRitem.techHours = qtechhrs;
                                Qbl.estHrslist.Add(HRitem);
                            }                           
                        }
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_blockadds";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.
                        string QName = sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                        string Type = sqlRd.GetString(sqlRd.GetOrdinal("Type"));
                        int Packidentifier = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                        string Itemname = sqlRd.GetString(sqlRd.GetOrdinal("ItemName"));
                        string Amount = sqlRd.GetString(sqlRd.GetOrdinal("Ammount"));
                        string descr = sqlRd.GetString(sqlRd.GetOrdinal("Descr"));
                        string cost = sqlRd.GetString(sqlRd.GetOrdinal("Cost"));

                        foreach (Quote_Block Qbl in Q_bList)
                        {
                            if (Qbl.QName == QName)
                            {
                                Adds additem = new Adds();
                                additem.ItemName = Itemname;
                                additem.Desc = descr;
                                additem.Ammount = Amount;
                                additem.Cost = cost;
                                Qbl.qbAdditionslist.Add(additem);
                            }
                        }
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quoteoptions";
                sqlRd = sqlCmd.ExecuteReader();                
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //quotefs(Qname varchar(100), HrRate double, HrAddition double, Additions double, Date varchar(20), primary key(Qname), unique(Qname))";
                        //I would also check for DB.Null here before reading the value.
                        foreach(Quote_Block Qitem in Q_bList)
                        { 
                            if(Qitem.QName == sqlRd.GetString(sqlRd.GetOrdinal("Qname")))
                            {
                                //qbitem.QName= sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                                Qitem.Qinfo.QuoteName = sqlRd.GetString(sqlRd.GetOrdinal("Qname"));
                                Qitem.Qinfo.Jobnum = sqlRd.GetString(sqlRd.GetOrdinal("jobnum"));
                                Qitem.Qinfo.JobDate = sqlRd.GetString(sqlRd.GetOrdinal("jobDate"));
                                Qitem.Qinfo.CompanyName = sqlRd.GetString(sqlRd.GetOrdinal("companyName"));
                                Qitem.Qinfo.Email = sqlRd.GetString(sqlRd.GetOrdinal("email"));
                                Qitem.Qinfo.Address = sqlRd.GetString(sqlRd.GetOrdinal("Address"));
                                Qitem.Qinfo.Suburb = sqlRd.GetString(sqlRd.GetOrdinal("Suburb"));
                                Qitem.Qinfo.State = sqlRd.GetString(sqlRd.GetOrdinal("State"));
                                Qitem.Qinfo.Postcode = sqlRd.GetString(sqlRd.GetOrdinal("PostCode"));
                                Qitem.Qinfo.ClientName = sqlRd.GetString(sqlRd.GetOrdinal("clientName"));
                                Qitem.Qinfo.ClientNum = sqlRd.GetString(sqlRd.GetOrdinal("clientNum"));
                                Qitem.Qinfo.Project = sqlRd.GetString(sqlRd.GetOrdinal("project"));
                                Qitem.Qinfo.TotalEst = sqlRd.GetString(sqlRd.GetOrdinal("totalEst"));
                                Qitem.Qinfo.Q1Allowance = sqlRd.GetString(sqlRd.GetOrdinal("Q1Allowance"));
                                Qitem.Qinfo.Q1Lm = sqlRd.GetString(sqlRd.GetOrdinal("Q1Lm"));
                                Qitem.Qinfo.Q1Days = sqlRd.GetString(sqlRd.GetOrdinal("Q1Days"));
                                Qitem.Qinfo.Q1Heading = sqlRd.GetString(sqlRd.GetOrdinal("Q1Heading"));
                                Qitem.Qinfo.Q1SlipRate = sqlRd.GetString(sqlRd.GetOrdinal("Q1SlipRate"));
                                Qitem.Qinfo.Q1StageVal = sqlRd.GetString(sqlRd.GetOrdinal("Q1StageVal"));
                                Qitem.Qinfo.Q1Warrenty = sqlRd.GetString(sqlRd.GetOrdinal("Q1Warrenty"));
                                Qitem.Qinfo.Q2Enable = sqlRd.GetBoolean(sqlRd.GetOrdinal("Q2Enable"));
                                Qitem.Qinfo.Q2Allowance = sqlRd.GetString(sqlRd.GetOrdinal("Q2Allowance"));
                                Qitem.Qinfo.Q2Lm = sqlRd.GetString(sqlRd.GetOrdinal("Q2Lm"));
                                Qitem.Qinfo.Q2Days = sqlRd.GetString(sqlRd.GetOrdinal("Q2Days"));
                                Qitem.Qinfo.Q2Heading = sqlRd.GetString(sqlRd.GetOrdinal("Q2Heading"));
                                Qitem.Qinfo.Q2SlipRate = sqlRd.GetString(sqlRd.GetOrdinal("Q2SlipRate"));
                                Qitem.Qinfo.Q2StageVal = sqlRd.GetString(sqlRd.GetOrdinal("Q2StageVal"));
                                Qitem.Qinfo.Q2Warrenty = sqlRd.GetString(sqlRd.GetOrdinal("Q2Warrenty"));
                                Qitem.Qinfo.Q3Enable = sqlRd.GetBoolean(sqlRd.GetOrdinal("Q3Enable"));
                                Qitem.Qinfo.Q3Allowance = sqlRd.GetString(sqlRd.GetOrdinal("Q3Allowance"));
                                Qitem.Qinfo.Q3Lm = sqlRd.GetString(sqlRd.GetOrdinal("Q3Lm"));
                                Qitem.Qinfo.Q3Days = sqlRd.GetString(sqlRd.GetOrdinal("Q3Days"));
                                Qitem.Qinfo.Q3Heading = sqlRd.GetString(sqlRd.GetOrdinal("Q3Heading"));
                                Qitem.Qinfo.Q3SlipRate = sqlRd.GetString(sqlRd.GetOrdinal("Q3SlipRate"));
                                Qitem.Qinfo.Q3StageVal = sqlRd.GetString(sqlRd.GetOrdinal("Q3StageVal"));
                                Qitem.Qinfo.Q3Warrenty = sqlRd.GetString(sqlRd.GetOrdinal("Q3Warrenty"));
                                Qitem.Qinfo.Rubbish = sqlRd.GetBoolean(sqlRd.GetOrdinal("Rubbish"));
                                Qitem.Qinfo.Phase3 = sqlRd.GetBoolean(sqlRd.GetOrdinal("Phase3"));
                                Qitem.Qinfo.Drytime = sqlRd.GetBoolean(sqlRd.GetOrdinal("Drytime"));
                                Qitem.Qinfo.JointColour = sqlRd.GetBoolean(sqlRd.GetOrdinal("JointColour"));
                                Qitem.Qinfo.Colour = sqlRd.GetBoolean(sqlRd.GetOrdinal("Colour"));
                                Qitem.Qinfo.SlipRes = sqlRd.GetBoolean(sqlRd.GetOrdinal("SlipRes"));
                                Qitem.Qinfo.QuoteStages = sqlRd.GetBoolean(sqlRd.GetOrdinal("QuoteStages"));
                                Qitem.Qinfo.ArrowHead = sqlRd.GetBoolean(sqlRd.GetOrdinal("ArrowHead"));
                                Qitem.Qinfo.AllowCtrlJoins = sqlRd.GetBoolean(sqlRd.GetOrdinal("AllowCtrlJoins"));
                                Qitem.Qinfo.JointCstr = sqlRd.GetString(sqlRd.GetOrdinal("jointCstr"));
                                Qitem.Qinfo.Colstr = sqlRd.GetString(sqlRd.GetOrdinal("Colourstr"));
                                Qitem.Qinfo.SlipRstr = sqlRd.GetString(sqlRd.GetOrdinal("slipR"));
                                Qitem.Qinfo.QuoteStagesstr = sqlRd.GetString(sqlRd.GetOrdinal("QuoteStagesStr"));
                                Qitem.Qinfo.Arrow = sqlRd.GetString(sqlRd.GetOrdinal("ArrowH"));
                                Qitem.Qinfo.AllowCtrlstr = sqlRd.GetString(sqlRd.GetOrdinal("AllowCtrl"));
                                //Qitem.Qinfo.Add(qbitem);
                            }
                        }
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quotecomms";
                sqlRd = sqlCmd.ExecuteReader();
                
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //quotefs(Qname varchar(100), HrRate double, HrAddition double, Additions double, Date varchar(20), primary key(Qname), unique(Qname))";
                        //I would also check for DB.Null here before reading the value.
                        foreach(Quote_Block Qitem in Q_bList)
                        {
                            if(Qitem.QName == sqlRd.GetString(sqlRd.GetOrdinal("Qname")))
                            {
                                if(sqlRd.GetString(sqlRd.GetOrdinal("Listtype")) == "general")
                                {
                                    string gcommentstr = sqlRd.GetString(sqlRd.GetOrdinal("Comment"));
                                    Qitem.Qinfo.CommentList.Add(gcommentstr);

                                }
                                if (sqlRd.GetString(sqlRd.GetOrdinal("Listtype")) == "Q1")
                                {
                                    string Q1commentstr = sqlRd.GetString(sqlRd.GetOrdinal("Comment"));
                                    string Q1Headingstr = sqlRd.GetString(sqlRd.GetOrdinal("Heading"));

                                    bool inList = false;
                                    foreach (QuoteComment Qcom in Qitem.Qinfo.Q1CommentList)
                                    {
                                        if(Qcom.QcHeading == Q1Headingstr)
                                        {
                                            inList = true;
                                            Qcom.QcTextList.Add(Q1commentstr);
                                        }
                                    }
                                    if(!inList)
                                    {
                                        QuoteComment Qcom = new QuoteComment();
                                        Qcom.QcHeading = Q1Headingstr;
                                        Qcom.QcTextList.Add(Q1commentstr);
                                        Qitem.Qinfo.Q1CommentList.Add(Qcom);
                                    }
                                }
                            }
                        }                        
                    }
                }
                sqlRd.Close();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void DelQuote(string Qimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_blockdata where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_blockadds where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_hrsdata where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_blockfs where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void AddtoQuote_BlockList(Quote_Block Qimport)
        {
            bool Inlist = false;
            foreach (Quote_Block qbItem in Q_bList)
            {
                if (Qimport.QName == qbItem.QName)
                {
                    Inlist = true;
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.quote_blockfs set Qname = @name, HrRate = @hrrate," +
                            "HrAddition = @hradds,Additions = @adds,Markup = @markup,Date = @date where Qname = @name";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@name", Qimport.QName);
                        sqlCmd.Parameters.AddWithValue("@hrrate", Qimport.savedHrRate);
                        sqlCmd.Parameters.AddWithValue("@hradds", Qimport.savedHradd);
                        sqlCmd.Parameters.AddWithValue("@adds", Qimport.savedAdds);
                        sqlCmd.Parameters.AddWithValue("@date", Qimport.SaveDate);
                        sqlCmd.Parameters.AddWithValue("@markup", Qimport.savedMarkup);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                    AddtoQuote_BlockData(Qimport);
                }
            }
            if (!Inlist)
            {
                // sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.quote_blockfs (Qname, HrRate, HrAddition, Additions,Markup, Date) values('" + Qimport.QName + "','" + Qimport.savedHrRate + "','" + Qimport.savedHradd +
                                        "','" + Qimport.savedAdds + "','" + Qimport.savedMarkup + "','" + Qimport.SaveDate + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                AddtoQuote_BlockData(Qimport);
            }
        }

        public void AddtoQuote_BlockData(Quote_Block Qimport)
        {
            // find last id in table
            
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_blockdata where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport.QName);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_blockadds where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport.QName);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.quote_hrsdata where QName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Qimport.QName);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }

            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_blockadds";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        ID3 = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                    }
                }
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_blockdata";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        ID = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                    }
                }
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.quote_hrsdata";
                sqlRd = sqlCmd.ExecuteReader();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        ID2 = sqlRd.GetInt32(sqlRd.GetOrdinal("PackId"));
                    }
                }
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            foreach (CSstage csitem in Qimport.qbStagelist)
            {
                try
                {
                    sqlConn.Open();
                    ID++;
                    sqlQuery = "insert into qepoxy2.quote_blockdata values('" + Qimport.QName + "','stage','" + ID + "','" + csitem.inputname + "','" + csitem.Ammount + "','"
                                    + csitem.Area + "','" + csitem.Lm + "','" + csitem.cost + "','" + csitem.packsAmnt + "','" + csitem.PackSystemName + "','" + csitem.Packvol + "','"
                                    + csitem.price + "','" + csitem.Staging + "','" + csitem.symbol + "','" + csitem.Systems + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }                
            }

            foreach (EstHoursSave estitem in Qimport.estHrslist)
            {
                try
                {
                    sqlConn.Open();
                    ID2++;
                    sqlQuery = "insert into qepoxy2.quote_hrsdata values('" + Qimport.QName + "','hours','" + ID2 + "','" + estitem.Hours + "','" + estitem.lm + "','"
                                    + estitem.m2 + "','" + estitem.method + "','" + estitem.stageName + "','" + estitem.systemName + "','" + estitem.techHours + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }

            foreach (Adds estitem in Qimport.qbAdditionslist)
            {
                try
                {
                    sqlConn.Open();
                    ID3++;
                    sqlQuery = "insert into qepoxy2.quote_blockadds values('" + Qimport.QName + "','Adds','" + ID3 + "','" + estitem.ItemName + "','" + estitem.Desc + "','"
                                    + estitem.Ammount + "','" + estitem.Cost + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }
        }

        public void AddtoQuoteGenLists(QuoteBlock Qbl)
        {
            foreach (Quote_Block qitem in Q_bList)
            {
                if (Qbl.QName == qitem.QName)
                {
                    qitem.Qinfo.Q2CommentList.Clear();
                    qitem.Qinfo.Q1CommentList.Clear();
                    qitem.Qinfo.Q3CommentList.Clear();
                    qitem.Qinfo.CommentList.Clear();

                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "delete from qepoxy2.quotecomms where QName = @Name";
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Qbl.QName);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }

                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "delete from qepoxy2.quoteoptions where QName = @Name";
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", Qbl.QName);
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add("-> " + msg);
                        sqlConn.Close();
                    }
                }
                break;
            }
            foreach (string qinfostr in Qbl.Qinfo.CommentList)
            {
                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','general','','','" + qinfostr + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    //sqlCmd.ExecuteReader();
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }   
            }
            foreach (QuoteComment Qcomm in Qbl.Qinfo.Q1CommentList)
            {
                foreach (string str in Qcomm.QcTextList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q1','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                }
            }
            foreach (QuoteComment Qcomm in Qbl.Qinfo.Q2CommentList)
            {
                foreach (string str in Qcomm.QcTextList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q2','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                }
            }
            foreach (QuoteComment Qcomm in Qbl.Qinfo.Q3CommentList)
            {
                foreach (string str in Qcomm.QcTextList)
                {
                    try
                    {
                        sqlConn.Open();
                        sqlQuery = "insert into qepoxy2.quotecomms values('" + Qbl.QName + "','Q3','" + Qcomm.QcHeading + "','" + Qcomm.QcSubHeading + "','" + str + "');";
                        sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                        //sqlCmd.ExecuteReader();
                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }
                }
            }
            try
            {
                sqlConn.Open();
                sqlQuery = "insert into qepoxy2.quoteoptions values( '" + Qbl.QName + "','" + Qbl.Qinfo.Jobnum + "','" + Qbl.Qinfo.JobDate + "','" + Qbl.Qinfo.CompanyName + "','" + Qbl.Qinfo.Email + "','" + Qbl.Qinfo.Address + "','" + Qbl.Qinfo.Suburb + "','" + Qbl.Qinfo.State + "','" + Qbl.Qinfo.Postcode + "','" + Qbl.Qinfo.ClientName + "','" + Qbl.Qinfo.ClientNum + "','" + Qbl.Qinfo.Project + "','" + Qbl.Qinfo.TotalEst + "','" +
                                                                    Qbl.Qinfo.Q1Allowance + "','" + Qbl.Qinfo.Q1Lm + "','" + Qbl.Qinfo.Q1Days + "','" + Qbl.Qinfo.Q1Heading + "','" + Qbl.Qinfo.Q1SlipRate + "','" + Qbl.Qinfo.Q1StageVal + "','" + Qbl.Qinfo.Q1Warrenty + "','" +
                                        Qbl.Qinfo.Q2Enable + "','" + Qbl.Qinfo.Q2Allowance + "','" + Qbl.Qinfo.Q2Lm + "','" + Qbl.Qinfo.Q2Days + "','" + Qbl.Qinfo.Q2Heading + "','" + Qbl.Qinfo.Q2SlipRate + "','" + Qbl.Qinfo.Q2StageVal + "','" + Qbl.Qinfo.Q2Warrenty + "','" +
                                        Qbl.Qinfo.Q3Enable + "','" + Qbl.Qinfo.Q3Allowance + "','" + Qbl.Qinfo.Q3Lm + "','" + Qbl.Qinfo.Q3Days + "','" + Qbl.Qinfo.Q3Heading + "','" + Qbl.Qinfo.Q3SlipRate + "','" + Qbl.Qinfo.Q3StageVal + "','" + Qbl.Qinfo.Q3Warrenty + "','" +
                                        Qbl.Qinfo.Rubbish + "','" + Qbl.Qinfo.Phase3 + "','" + Qbl.Qinfo.Drytime + "','" + Qbl.Qinfo.JointColour + "','" + Qbl.Qinfo.Colour + "','" + Qbl.Qinfo.SlipRes + "','" + Qbl.Qinfo.QuoteStages + "','" + Qbl.Qinfo.ArrowHead + "','" + Qbl.Qinfo.AllowCtrlJoins + "','" +
                                        Qbl.Qinfo.JointCstr + "','" + Qbl.Qinfo.Colstr + "','" + Qbl.Qinfo.SlipRstr + "','" + Qbl.Qinfo.QuoteStagesstr + "','" + Qbl.Qinfo.Arrow + "','" + Qbl.Qinfo.AllowCtrlstr + "');";

                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                //sqlCmd.ExecuteReader();
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }                
            
        }

        public void UpdateHrrates(Hrrates indata)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.hrratesettings where HrName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", indata.hrsname);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                sqlQuery = "insert into qepoxy2.hrratesettings (HrName, HrWeight) values('" + indata.hrsname + "','" + indata.hrsWeight + "');";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlRd = sqlCmd.ExecuteReader();
                sqlConn.Close();

            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }

        public void UpdateAdditions(string indata)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.additions where ItemName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", indata);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
            try
            {
                sqlConn.Open();
                sqlQuery = "insert into qepoxy2.additions (ItemName) values('" + indata + "');";
                sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                sqlRd = sqlCmd.ExecuteReader();
                sqlConn.Close();

            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }


        public void AddtoHrrates(Hrrates indata)
        {
            bool Inlist = false;
            foreach (Hrrates HrItem in HrrateList)
            {
                if (indata.hrsname == HrItem.hrsname)
                {
                    Inlist = true;
                    //sqlConn.ConnectionString = "server=" + server + ";" + "user id=" + username + ";" + "password=" + password + ";" + "database=" + database;
                    try
                    {
                        sqlConn.Open();
                        MySqlCommand sqlCmd = new MySqlCommand();
                        sqlCmd.Connection = sqlConn;
                        sqlCmd.CommandText = "Update qepoxy2.hrratesettings set HrName = @Name, HrWeight = @Weight";

                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@Name", indata.hrsname);
                        sqlCmd.Parameters.AddWithValue("@Weight", indata.hrsWeight);



                        sqlCmd.ExecuteNonQuery();
                        sqlConn.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;
                        SQLErrorList.Add(msg);
                        sqlConn.Close();
                    }

                }
            }
            if (!Inlist)
            {
                try
                {
                    sqlConn.Open();
                    MySqlCommand sqlCmd = new MySqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = "delete from qepoxy2.hrratesettings where HrName = @Name";
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@Name", indata.hrsname);
                    sqlCmd.ExecuteNonQuery();
                    sqlConn.Close();
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }

                try
                {
                    sqlConn.Open();
                    sqlQuery = "insert into qepoxy2.hrratesettings (HrName, HrWeight) values('" + indata.hrsname + "','" + indata.hrsWeight + "');";
                    sqlCmd = new MySqlCommand(sqlQuery, sqlConn);
                    sqlRd = sqlCmd.ExecuteReader();
                    sqlConn.Close();

                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    SQLErrorList.Add(msg);
                    sqlConn.Close();
                }
            }
            RebuildHrRates();

        }
        public void DelItemFromHrrates(string Pimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.hrratesettings where HrName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Pimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void DelItemFromAdditions(string Pimport)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.Additions where ItemName = @Name";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Name", Pimport);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void RebuildHrRates()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.hrratesettings";
                sqlRd = sqlCmd.ExecuteReader();
                HrrateList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        //I would also check for DB.Null here before reading the value.                        
                        Hrrates HrItem = new Hrrates();
                        HrItem.hrsname = sqlRd.GetString(sqlRd.GetOrdinal("HrName"));
                        HrItem.hrsWeight = sqlRd.GetDouble(sqlRd.GetOrdinal("HrWeight"));
                        HrrateList.Add(HrItem);
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }

        public void RebuildAdditions()
        {
            try
            {
                sqlConn.Open();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "SELECT * FROM qepoxy2.Additions";
                sqlRd = sqlCmd.ExecuteReader();
                AdditionsList.Clear();
                if (sqlRd.HasRows)
                {
                    while (sqlRd.Read())
                    {
                        string strItem = sqlRd.GetString(sqlRd.GetOrdinal("ItemName"));
                        AdditionsList.Add(strItem);
                    }
                }
                sqlConn.Close();
                sqlRd.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
                sqlRd.Close();
            }
        }
        public void DelItemFromHrRates(string Instr)
        {
            try
            {
                sqlConn.Open();
                MySqlCommand sqlCmd = new MySqlCommand();
                sqlCmd.Connection = sqlConn;
                sqlCmd.CommandText = "delete from qepoxy2.hrratesettings where Comment = @Comm";
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@Comm", Instr);
                sqlCmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                SQLErrorList.Add(msg);
                sqlConn.Close();
            }
        }
        public void Addtolog(string inputstr)
        {
            File.AppendAllText(@"..\systemLog.log", DateTime.Now.ToString("hh:mm:ss ") + " " + inputstr + Environment.NewLine);
        }
        public void DataSaveSystemsData(string SaveLocation)
        {
            // --------------------------------------------------------------------------------------
            // old save proceedure
            // --------------------------------------------------------------------------------------
            string filepath = SaveLocation;
            System.IO.Directory.CreateDirectory(filepath);
            string pathString = System.IO.Path.Combine(filepath, "FileNames.csv");
            //string curFile = @filepath+"FileNames.csv";

            System.IO.File.WriteAllBytes(pathString, new byte[0]);

            string titleStr = "Name,Type";
            string[] outArray = new string[SystemsList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;

            foreach (ProductSystem psys in SystemsList)
            {
                outArray[i] = psys.BaseMat.ToString() + "," + psys.Type.ToString();
                //Addtolog(outArray[i]);
                string FileName = psys.BaseMat + ".csv";
                string pathString2 = System.IO.Path.Combine(filepath, FileName);
                
                string curFile2 = @pathString2;

                System.IO.File.WriteAllBytes(curFile2, new byte[0]);
                string titleStr1 = "Type,Name,Price,Weight,Volume,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
                List<string> outList = new List<string>();

                foreach (Pack pItem in psys.PackList)
                {
                    string pdout = "Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType;
                    outList.Add(pdout);

                    foreach (Priceitem prItem in pItem.PriceItemList)
                    {
                        string plOut = "sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                     prItem.Volume + "," + prItem.Coverage + "," + prItem.Unit+","+prItem.productDate;
                        outList.Add(plOut);

                    }
                }

                string[] outArray1 = new string[outList.Count() + 1];
                outArray1[0] = titleStr1;
                int j = 1;
                foreach (string stritem in outList)
                {
                    outArray1[j] = stritem;
                    j++;
                }
                System.IO.File.WriteAllLines(@filepath + FileName, outArray1);
                i++;
            }
            System.IO.File.WriteAllLines(@filepath+ "FileNames.csv", outArray);
            Addtolog("write to file FileNames.csv");
        }
        public void WriteNewSysFile(string SaveLocation,string FileName,List<Pack> pList)
        {
            string filepath = SaveLocation;
            System.IO.Directory.CreateDirectory(filepath);
            string pathString = System.IO.Path.Combine(filepath, FileName);
            string curFile2 = @pathString+".csv";
            System.IO.File.WriteAllBytes(curFile2, new byte[0]);

            string titleStr1 = "Type,Name,Price,Weight,Volume,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();

            foreach (Pack pItem in pList)
            {
                string pdout = "Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                            pItem.Volume + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                            pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType;
                outList.Add(pdout);


                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    string plOut = "sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate;
                    outList.Add(plOut);
                }
            }

            string[] outArray1 = new string[outList.Count() + 1];
            outArray1[0] = titleStr1;

            int j = 1;
            foreach (string stritem in outList)
            {
                outArray1[j] = stritem;

                j++;
            }            
            System.IO.File.WriteAllLines(@pathString + ".csv", outArray1);
        }
        public void DataSavePriceItemsData(string filepath)
        {
            string FileName = "Products.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(@"pathString", new byte[0]);
            string titleStr = "Name,Price,Weight,Volume,Coverage,Unit,ProductDate";
            string[] outArray = new string[PriceList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;

            foreach (Priceitem Pi in PriceList)
            {
                outArray[i] = Pi.Name + "," + Pi.Price + "," + Pi.Weight + "," + Pi.Volume + "," + Pi.Coverage + "," + Pi.Unit + "," + Pi.productDate;
                i++;
            }
            System.IO.File.WriteAllLines(@pathString, outArray);

        }

        public void DataSaveAdditions(string filepath)
        {
            string FileName = "Additions.txt";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(@"pathString", new byte[0]);
            string titleStr = "ItemName";
            string[] outArray = new string[AdditionsList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;

            foreach (string strItem in AdditionsList)
            {
                outArray[i] = strItem;
                i++;
            }
            System.IO.File.WriteAllLines(@pathString, outArray);
        }
        public void LoadAdditions(string filepath)
        {
            List<string> OutList = new List<string>();
            string FileName = "Additions.txt";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString);

            // string[] lines = System.IO.File.ReadAllLines(@"..\..\Data\\\HourRates.csv");
            for (int i = 1; i < lines.Length; i++)
            {
                if (lines[i] != "")
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    
                    string item = parts[0];                   
                    AdditionsList.Add(item);
                }
            }           
        }

        public void DataSaveHrRates(string filepath)
        {
            string FileName = "HourRates.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(@"pathString", new byte[0]);
            string titleStr = "HrName,HrWeight";
            string[] outArray = new string[HrrateList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;

            foreach (Hrrates HrItem in HrrateList)
            {
                outArray[i] = HrItem.hrsname + "," + HrItem.hrsWeight;
                i++;
            }
            System.IO.File.WriteAllLines(@pathString, outArray);
        }
        public void LoadHrRates(string filepath)
        {
            List<ProductSystem> OutPsList = new List<ProductSystem>();
            string FileName = "HourRates.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString);

           // string[] lines = System.IO.File.ReadAllLines(@"..\..\Data\\\HourRates.csv");
            for (int i = 1; i < lines.Length; i++)
            {
                if(lines[i] != "")
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    Hrrates hrrateItem = new Hrrates();
                    hrrateItem.hrsname = parts[0];
                    hrrateItem.hrsWeight = double.Parse(parts[1]);
                    HrrateList.Add(hrrateItem);
                    //AddtoHrrates(hrrateItem);
                }
            }
            //RebuildHrRates();            
        }
        public List<ProductSystem> LoadFSData(string filepath)
        {
             List<ProductSystem> OutPsList = new List<ProductSystem>();
            string FileName = "FileNames.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString);
            for (int i = 1; i < lines.Length; i++)
            {
                ProductSystem Psys = new ProductSystem();
                var tmpValue = lines[i];
                var parts = tmpValue.Split(',');
                string Infs = parts[0].ToString();
                Psys.BaseMat = Infs;
                Psys.Type = parts[1];
                Psys.PackList = LoadSystemsData(Infs, filepath);
                OutPsList.Add(Psys);
            }
            return OutPsList;
        }
        public void DataSavePackData(string filepath)
        {
            //------------------------------------------------------------------------------------------------
            // old save proceedure
            //------------------------------------------------------------------------------------------------
            string FileName = "PackData.csv";            
            string pathString = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(@pathString, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();
            foreach (Pack pItem in PackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Coverage + "," + pItem.Unit);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," + prItem.Volume + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString, outArray);
        }
        public List<Pack> LoadPackData(string filepath)
        {
            Pack nP = new Pack();
            List<Pack> OutPackList = new List<Pack>();
            string FileName = "PackData.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString);
            for (int i = 1; i < lines.Length; i++)
            {
                var tmpValue = lines[i];
                var parts = tmpValue.Split(',');
                if (parts[0] == "Pack")
                {
                    nP.Name = parts[1];
                    nP.Price = double.Parse(parts[2]);
                    nP.Weight = double.Parse(parts[3]);
                    nP.Volume = double.Parse(parts[4]);
                    nP.Coverage = double.Parse(parts[5]);
                    nP.Unit = parts[6];
                    nP.Meterage = 0;

                    if (i + 1 < lines.Length)
                    {
                        var NexttmpValue = lines[i + 1];
                        var Nextparts = NexttmpValue.Split(',');
                        if (Nextparts[0] == "Pack")
                        {
                            bool itemInList = false;
                            foreach (Pack pItem in OutPackList)
                            {
                                if (nP.Name == pItem.Name)
                                {
                                    itemInList = true;
                                }
                            }
                            if (!itemInList)
                            {
                                OutPackList.Add(nP);
                            }
                            nP = new Pack();
                        }
                    }
                    else if (i + 1 == lines.Length)
                    {
                        bool itemInList = false;
                        foreach (Pack pItem in OutPackList)
                        {
                            if (nP.Name == pItem.Name)
                            {
                                itemInList = true;
                            }
                        }
                        if (!itemInList)
                        {
                            OutPackList.Add(nP);
                        }
                    }
                }
                if (parts[0] == "sub")
                {
                    Priceitem Pi = new Priceitem();
                    Pi.Name = parts[1];
                    Pi.Price = double.Parse(parts[2]);
                    Pi.Weight = double.Parse(parts[3]);
                    Pi.Volume = double.Parse(parts[4]);
                    Pi.Coverage = double.Parse(parts[5]);
                    Pi.Unit = parts[6];
                    Pi.Meterage = 0;
                    //Pi.productDate = parts[7];
                    Pi.productDate = "";
                    nP.AddChild(Pi);
                    if (i + 1 < lines.Length)
                    {
                        var NexttmpValue = lines[i + 1];
                        var Nextparts = NexttmpValue.Split(',');
                        if (Nextparts[0] == "Pack")
                        {
                            bool itemInList = false;
                            foreach (Pack pItem in OutPackList)
                            {
                                if (nP.Name == pItem.Name)
                                {
                                    itemInList = true;
                                }
                            }
                            if (!itemInList)
                            {
                                OutPackList.Add(nP);
                            }
                            nP = new Pack();
                        }
                    }
                    else if (i + 1 == lines.Length)
                    {
                        bool itemInList = false;
                        foreach (Pack pItem in OutPackList)
                        {
                            if (nP.Name == pItem.Name)
                            {
                                itemInList = true;
                            }
                        }
                        if (!itemInList)
                        {
                            OutPackList.Add(nP);
                        }
                    }
                }
            }
            return OutPackList;
        }
        public List<Pack> LoadSystemsData(string FloorSystem,string filepath)
        {
            List<Pack> OutList = new List<Pack>();
            Pack nP = new Pack();
            string FileName = FloorSystem + ".csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            if (File.Exists(@pathString)) { 
                string[] lines = System.IO.File.ReadAllLines(@pathString);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Coverage = double.Parse(parts[5]);
                        nP.Unit = parts[6];
                        nP.Options.Grind = parts[7];
                        nP.Options.RollCoat = parts[8];
                        nP.Options.Trowel = parts[9];
                        nP.Options.Polivac = int.Parse(parts[10]);
                        nP.Options.CovingType = parts[11];
                        nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                //AddPackToList(nP);
                                OutList.Add(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            //AddPackToList(nP);
                            OutList.Add(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        Pi.Coverage = double.Parse(parts[5]);
                        Pi.Unit = parts[6];
                        Pi.Meterage = 0;
                        //Pi.productDate = parts[7];
                        Pi.productDate = "";
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                //AddPackToList(nP);
                                OutList.Add(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            //AddPackToList(nP);
                            OutList.Add(nP);
                        }
                    }
                }
            }
            return OutList;
        }
        public List<Priceitem> LoadProductData(string filepath)
        {
            List<Priceitem> OutPriceList = new List<Priceitem>();
            string FileName = "Products.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);

            string[] lines = System.IO.File.ReadAllLines(@pathString);
            for (int i = 1; i < lines.Length; i++)
            {
                Priceitem Pitem = new Priceitem();
                var tmpValue = lines[i];
                var parts = tmpValue.Split(',');

                Pitem.Name = parts[0];
                Pitem.Price = double.Parse(parts[1]);
                Pitem.Weight = double.Parse(parts[2]);
                Pitem.Volume = double.Parse(parts[3]);
                Pitem.Coverage = double.Parse(parts[4]);
                Pitem.Unit = parts[5];
                Pitem.productDate = parts[6];
                //Pitem.productDate = "";
                OutPriceList.Add(Pitem);
            }
            return OutPriceList;
        }
        public void ClearLists()
        {
            PackList.Clear();
            SystemsList.Clear();

            PriceList.Clear();
        }
        public void AddPackToList(Pack newPack,List<Pack> Outlist)
        {
            bool itemInList = false;           
            foreach (Pack pItem in PackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                Outlist.Add(newPack);                
            }
        }

        public List<Quote_Block> LoadQuoteFSdata(string filepath)
        {
            //string Gc;
            List<Quote_Block> OutQbList = new List<Quote_Block>();
            string FileName = "QuoteFs.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString2);
            for (int i = 1; i < lines.Length; i++)
            {
                string tmpValue = lines[i];
                var parts = tmpValue.Split(',');

                Quote_Block LocalQb = new Quote_Block();
                LocalQb.QName = parts[0];
                LocalQb.savedHrRate = parts[1];
                LocalQb.savedHradd = parts[2];
                LocalQb.savedAdds = parts[3];
                LocalQb.savedMarkup = parts[4];
                LocalQb.SaveDate = parts[5];
                LocalQb.LoadQuoteHrsList(filepath);
                LocalQb.LoadQuotestageList(filepath);
                OutQbList.Add(LocalQb);
            }
            OutQbList.Sort((x, y) => x.QName.CompareTo(y.QName));
            return OutQbList;
        }
        public void DataSaveQuoteFSdata(string filepath)
        {
            string FileName = "QuoteFs.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.Directory.CreateDirectory(path);
            System.IO.File.WriteAllBytes(pathString2, new byte[0]);
            string titleStr = "QuoteFSName,hrrate,HrAdditions,Additions,Markup,Date";
            string[] outArray = new string[Q_bList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;
            foreach (Quote_Block Qitem in Q_bList)
            {
                outArray[i] = Qitem.QName +","+ Qitem.savedHrRate + "," +Qitem.savedHradd + "," + Qitem.savedAdds + "," + Qitem.savedMarkup + "," + Qitem.SaveDate;
                i++;
            }
            System.IO.File.WriteAllLines(pathString2, outArray);
        }
        public void dataSaveCompanyData(string filepath)
        {
            string FileName = "Companies.csv";
            string pathString = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(pathString, new byte[0]);
            string titleStr = "Name,Email,Address,Suburb,State,ContantName,ContactNum";
            string[] outArray = new string[Complist.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;

            foreach (Company Citem in Complist)
            {

                outArray[i] = Citem.CompName + "," + Citem.Email + "," + Citem.Address+ ","+ Citem.Suburb + "," + Citem.State + "," + Citem.ContactName + "," + Citem.ContactNum;
                i++;
            }
            System.IO.File.WriteAllLines(pathString, outArray);
        }
        public void LoadCompanyData(string filepath)
        {
            string FileName = "Companies.csv";
            string pathString2 = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString2); 
            for (int i = 1; i < lines.Length; i++)
            {
                Company citem = new Company();
                var tmpValue = lines[i];
                var parts = tmpValue.Split(',');

                citem.CompName = parts[0];
                citem.Email = parts[1];
                citem.Address = parts[2];
                citem.ContactName = parts[3];
                Complist.Add(citem);
            }
        }
        public void DatasaveQuoteData(string filepath)
        {
            foreach(Quote_Block qbin in Q_bList)
            {
                string FileName = qbin.QName + "-QuoteData.csv";
                string PTH2 = "Quotes";
                string path = System.IO.Path.Combine(filepath, PTH2);
                string pathString2 = System.IO.Path.Combine(path, FileName);

                qbin.SaveQuoteHrsList(filepath);
                qbin.SaveQuotestageList(filepath);
                //qbin.SaveFPackData(filepath);
                //qbin.SaveSFPackData(filepath);
                //qbin.SaveCPackData(filepath);
                //qbin.SaveCVPackData(filepath);
                //qbin.SaveLMPackData(filepath);
                //qbin.SavePCPackData(filepath);

                //string FileName = Namein + "-QuoteData";
                System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
                string titleStr = "Type,jobnum,jobDate,companyName,email,Address,Suburb,State,PostCode,clientName,ClientNum,project,totalEst," +
                    "Q1Allowance,Q1Lm,Q1Days,Q1Heading,Q1SlipRate,Q1StageVal,Q1Warrenty," +
                    "Q2Enable,Q2Allowance,Q2Lm,Q2Days,Q2Heading,Q2SlipRate,Q2StageVal,Q2Warrenty," +
                    "Q3Enable,Q3Allowance,Q3Lm,Q3Days,Q3Heading,Q3SlipRate,Q3StageVal,Q3Warrenty," +
                    "Rubbish,Phase3,Drytime,JointColour,Colour,SlipRes,QuoteStages,ArrowHead,AllowCtrlJoins,JointCstr,Colstr,SlipRstr,QuoteStagesstr,Arrow,AllowctrlJoinstr";

                List<string> outList = new List<string>();
                outList.Add(titleStr);
                outList.Add("data" + "," +
                qbin.Qinfo.Jobnum + "," +
                qbin.Qinfo.JobDate + "," +
                qbin.Qinfo.CompanyName + "," +
                qbin.Qinfo.Email + "," +
                qbin.Qinfo.Address + "," +
                qbin.Qinfo.Suburb + "," +
                qbin.Qinfo.State + "," +
                qbin.Qinfo.Postcode + "," +
                qbin.Qinfo.ClientName + "," +
                qbin.Qinfo.ClientNum + "," +
                qbin.Qinfo.Project + "," +
                qbin.Qinfo.TotalEst + "," +

                qbin.Qinfo.Q1Allowance + "," +
                qbin.Qinfo.Q1Lm + "," +
                qbin.Qinfo.Q1Days + "," +
                qbin.Qinfo.Q1Heading + "," +
                qbin.Qinfo.Q1SlipRate + "," +
                qbin.Qinfo.Q1StageVal + "," +
                qbin.Qinfo.Q1Warrenty + "," +

                qbin.Qinfo.Q2Enable + "," +
                qbin.Qinfo.Q2Allowance + "," +
                qbin.Qinfo.Q2Lm + "," +
                qbin.Qinfo.Q2Days + "," +
                qbin.Qinfo.Q2Heading + "," +
                qbin.Qinfo.Q2SlipRate + "," +
                qbin.Qinfo.Q2StageVal + "," +
                qbin.Qinfo.Q2Warrenty + "," +

                qbin.Qinfo.Q3Enable + "," +
                qbin.Qinfo.Q3Allowance + "," +
                qbin.Qinfo.Q3Lm + "," +
                qbin.Qinfo.Q3Days + "," +
                qbin.Qinfo.Q3Heading + "," +
                qbin.Qinfo.Q3SlipRate + "," +
                qbin.Qinfo.Q3StageVal + "," +
                qbin.Qinfo.Q3Warrenty + "," +

                qbin.Qinfo.Rubbish + "," +
                qbin.Qinfo.Phase3 + "," +
                qbin.Qinfo.Drytime + "," +
                qbin.Qinfo.JointColour + "," +
                qbin.Qinfo.Colour + "," +
                qbin.Qinfo.SlipRes + "," +
                qbin.Qinfo.QuoteStages + "," +
                qbin.Qinfo.ArrowHead + "," +
                qbin.Qinfo.AllowCtrlJoins + "," +

                qbin.Qinfo.JointCstr + "," +
                qbin.Qinfo.Colstr + "," +
                qbin.Qinfo.SlipRstr + "," +
                qbin.Qinfo.QuoteStagesstr + "," +
                qbin.Qinfo.Arrow + "," +
                qbin.Qinfo.AllowCtrlstr);

                foreach (QuoteComment Qitem in qbin.Qinfo.Q1CommentList)
                {
                    outList.Add("heading1," + Qitem.QcHeading);
                    foreach (string str in Qitem.QcTextList)
                    {
                        //string str2 = str.Replace(",", "^");
                        //outList.Add("text1," + str2);
                        outList.Add("text1," + str.Replace("^", ","));
                    }
                }
                foreach (QuoteComment Qitem in qbin.Qinfo.Q2CommentList)
                {
                    outList.Add("heading2," + Qitem.QcHeading);
                    foreach (string str in Qitem.QcTextList)
                    {
                        outList.Add("text2," + str.Replace("^", ","));
                    }
                }
                foreach (QuoteComment Qitem in qbin.Qinfo.Q3CommentList)
                {
                    outList.Add("heading3," + Qitem.QcHeading);
                    foreach (string str in Qitem.QcTextList)
                    {
                        outList.Add("text3," + str.Replace("^",","));
                    }
                }
                foreach (string str in qbin.Qinfo.CommentList)
                {
                    outList.Add("Gcom," + str.Replace("^", ","));
                }
                System.IO.File.WriteAllLines(@pathString2, outList);
            }            
        }
        public void LoadGCommentsDb(string filepath)
        {
            string FileName = "GeneralCommentsDb.txt";
            string pathString2 = System.IO.Path.Combine(filepath, FileName);
            string[] lines = System.IO.File.ReadAllLines(@pathString2);
            for (int i = 1; i < lines.Length; i++)
            {
                string tmpValue = lines[i];
                GCommentdbList.Add(tmpValue);
            }
        }
        public void DataSaveGCommentsDb(string filepath)
        {
            string FileName = "GeneralCommentsDb.txt";
            string pathString2 = System.IO.Path.Combine(filepath, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type_GComment";
            string[] outArray = new string[GCommentdbList.Count() + 1];
            outArray[0] = titleStr;
            int i = 1;
            foreach (string Gcitem in GCommentdbList)
            {
                outArray[i] = Gcitem;
                i++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void DataSaveDbToDisk(string filepath)
        {
            string FileName = "CommentDb.txt";
            string pathString2 = System.IO.Path.Combine(filepath, FileName);            
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type_Comment";
            List<string> outList = new List<string>();

            foreach (QuoteComment QcItm in QuoteDbList)
            {
                outList.Add("Heading" + "_" + QcItm.QcHeading);
                foreach (string StrItem in QcItm.QcTextList)
                {
                    outList.Add("Comment" + "_" + StrItem);
                }
            }
            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadDbFromDisk(string filepath)
        {
            string FileName = "CommentDb.txt";
            string pathString2 = System.IO.Path.Combine(filepath, FileName);
            if (File.Exists(@pathString2))
            {
                QuoteComment Qc = new QuoteComment();
                string[] lines = File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split('_');
                    if (parts[0] == "Heading")
                    {
                        Qc.QcHeading = parts[1];

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split('_');
                            if (Nextparts[0] == "Heading")
                            {
                                QuoteDbList.Add(Qc);
                                Qc = new QuoteComment();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            QuoteDbList.Add(Qc);
                        }
                    }
                    if (parts[0] == "Comment")
                    {
                        string strItem = parts[1];
                        Qc.QcTextList.Add(strItem);

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split('_');
                            if (Nextparts[0] == "Heading")
                            {
                                QuoteDbList.Add(Qc);
                                Qc = new QuoteComment();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            QuoteDbList.Add(Qc);
                        }
                    }
                }
            }
        }
        public void DataloadQuoteData(Quote_Block Qblimport,string filepath)
        {
            if (Qblimport.QName != "")
            {
                Quote_Block NewQbl = new Quote_Block();
                NewQbl = Qblimport;

                string FileName = NewQbl.QName + "-QuoteData.csv";
                string PTH2 = "Quotes";
                string path = System.IO.Path.Combine(filepath, PTH2);
                string pathString2 = System.IO.Path.Combine(path, FileName);

                //QupdateBox(FileName + " Data Has Been Loaded");
                if (!File.Exists(@pathString2))
                {
                    //QupdateBox(FileName + " Does Not Exist");
                }
                else
                {
                    QuoteComment qc1 = new QuoteComment();
                    QuoteComment qc2 = new QuoteComment();
                    QuoteComment qc3 = new QuoteComment();

                    NewQbl.Qinfo.Q1CommentList.Clear();
                    NewQbl.Qinfo.Q2CommentList.Clear();
                    NewQbl.Qinfo.Q3CommentList.Clear();
                    NewQbl.Qinfo.CommentList.Clear();

                    string[] lines = System.IO.File.ReadAllLines(@pathString2);
                    for (int i = 1; i < lines.Length; i++)
                    {
                        var tmpValue = lines[i];
                        var parts = tmpValue.Split(',');
                        if (parts[0] == "data")
                        {
                            if(!legacyData)
                            { 
                                NewQbl.Qinfo.Jobnum = parts[1];
                                NewQbl.Qinfo.JobDate = parts[2];
                                NewQbl.Qinfo.CompanyName = parts[3];
                                NewQbl.Qinfo.Email = parts[4];
                                NewQbl.Qinfo.Address = parts[5];
                                NewQbl.Qinfo.Suburb = parts[6];
                                NewQbl.Qinfo.State = parts[7];
                                NewQbl.Qinfo.Postcode = parts[8];
                                NewQbl.Qinfo.ClientName = parts[9];
                                NewQbl.Qinfo.ClientNum = parts[10];
                                NewQbl.Qinfo.Project = parts[11];
                                NewQbl.Qinfo.TotalEst = parts[12];

                                NewQbl.Qinfo.Q1Allowance = parts[13];
                                NewQbl.Qinfo.Q1Lm = parts[14];
                                NewQbl.Qinfo.Q1Days = parts[15];
                                NewQbl.Qinfo.Q1Heading = parts[16];
                                NewQbl.Qinfo.Q1SlipRate = parts[17];
                                NewQbl.Qinfo.Q1StageVal = parts[18];
                                NewQbl.Qinfo.Q1Warrenty = parts[19];

                                string boolcheck1 = parts[20].ToLower();
                                NewQbl.Qinfo.Q2Enable = bool.Parse(boolcheck1);
                                NewQbl.Qinfo.Q2Allowance = parts[21];
                                NewQbl.Qinfo.Q2Lm = parts[22];
                                NewQbl.Qinfo.Q2Days = parts[23];
                                NewQbl.Qinfo.Q2Heading = parts[24];
                                NewQbl.Qinfo.Q2SlipRate = parts[25];
                                NewQbl.Qinfo.Q2StageVal = parts[26];
                                NewQbl.Qinfo.Q2Warrenty = parts[27];

                                string boolcheck2 = parts[28].ToLower();
                                NewQbl.Qinfo.Q3Enable = bool.Parse(boolcheck2);
                                NewQbl.Qinfo.Q3Allowance = parts[29];
                                NewQbl.Qinfo.Q3Lm = parts[30];
                                NewQbl.Qinfo.Q3Days = parts[31];
                                NewQbl.Qinfo.Q3Heading = parts[32];
                                NewQbl.Qinfo.Q3SlipRate = parts[33];
                                NewQbl.Qinfo.Q3StageVal = parts[34];
                                NewQbl.Qinfo.Q3Warrenty = parts[35];

                                string boolcheck3 = parts[36].ToLower();
                                string boolcheck4 = parts[37].ToLower();
                                string boolcheck5 = parts[38].ToLower();
                                string boolcheck6 = parts[39].ToLower();
                                string boolcheck7 = parts[40].ToLower();
                                string boolcheck8 = parts[41].ToLower();
                                string boolcheck9 = parts[42].ToLower();
                                string boolcheck10 = parts[43].ToLower();
                                string boolcheck11 = parts[44].ToLower();

                                NewQbl.Qinfo.Rubbish = bool.Parse(boolcheck3);
                                NewQbl.Qinfo.Phase3 = bool.Parse(boolcheck4);
                                NewQbl.Qinfo.Drytime = bool.Parse(boolcheck5);
                                NewQbl.Qinfo.JointColour = bool.Parse(boolcheck6);
                                NewQbl.Qinfo.Colour = bool.Parse(boolcheck7);
                                NewQbl.Qinfo.SlipRes = bool.Parse(boolcheck8);
                                NewQbl.Qinfo.QuoteStages = bool.Parse(boolcheck9);
                                NewQbl.Qinfo.ArrowHead = bool.Parse(boolcheck10);
                                NewQbl.Qinfo.AllowCtrlJoins = bool.Parse(boolcheck11);

                                NewQbl.Qinfo.JointCstr = parts[45];
                                NewQbl.Qinfo.Colstr = parts[46];
                                NewQbl.Qinfo.SlipRstr = parts[47];
                                NewQbl.Qinfo.QuoteStagesstr = parts[48];
                                NewQbl.Qinfo.Arrow = parts[49];
                                NewQbl.Qinfo.JointCstr = parts[50];

                            }
                            else
                            {
                                //------------------------------------------------------------------------------
                                NewQbl.Qinfo.Jobnum = parts[1];
                                NewQbl.Qinfo.JobDate = parts[2];
                                NewQbl.Qinfo.CompanyName = parts[3];
                                NewQbl.Qinfo.Email = parts[4];
                                NewQbl.Qinfo.Address = parts[5];
                                NewQbl.Qinfo.ClientName = parts[6];
                                NewQbl.Qinfo.Project = parts[7];
                                NewQbl.Qinfo.TotalEst = parts[8];

                                NewQbl.Qinfo.Q1Allowance = parts[9];
                                NewQbl.Qinfo.Q1Lm = parts[10];
                                NewQbl.Qinfo.Q1Days = parts[11];
                                NewQbl.Qinfo.Q1Heading = parts[12];
                                NewQbl.Qinfo.Q1SlipRate = parts[13];
                                NewQbl.Qinfo.Q1StageVal = parts[14];
                                NewQbl.Qinfo.Q1Warrenty = parts[15];

                                string boolcheck1 = parts[16].ToLower();
                                NewQbl.Qinfo.Q2Enable = bool.Parse(boolcheck1);
                                NewQbl.Qinfo.Q2Allowance = parts[17];
                                NewQbl.Qinfo.Q2Lm = parts[18];
                                NewQbl.Qinfo.Q2Days = parts[19];
                                NewQbl.Qinfo.Q2Heading = parts[20];
                                NewQbl.Qinfo.Q2SlipRate = parts[21];
                                NewQbl.Qinfo.Q2StageVal = parts[22];
                                NewQbl.Qinfo.Q2Warrenty = parts[23];

                                string boolcheck2 = parts[24].ToLower();
                                NewQbl.Qinfo.Q3Enable = bool.Parse(boolcheck2);
                                NewQbl.Qinfo.Q3Allowance = parts[25];
                                NewQbl.Qinfo.Q3Lm = parts[26];
                                NewQbl.Qinfo.Q3Days = parts[27];
                                NewQbl.Qinfo.Q3Heading = parts[28];
                                NewQbl.Qinfo.Q3SlipRate = parts[29];
                                NewQbl.Qinfo.Q3StageVal = parts[30];
                                NewQbl.Qinfo.Q3Warrenty = parts[31];

                                string boolcheck3 = parts[32].ToLower();
                                string boolcheck4 = parts[33].ToLower();
                                string boolcheck5 = parts[34].ToLower();
                                string boolcheck6 = parts[35].ToLower();
                                string boolcheck7 = parts[36].ToLower();
                                string boolcheck8 = parts[37].ToLower();
                                string boolcheck9 = parts[38].ToLower();
                                string boolcheck10 = parts[39].ToLower();
                                string boolcheck11 = parts[40].ToLower();

                                NewQbl.Qinfo.Rubbish = bool.Parse(boolcheck3);
                                NewQbl.Qinfo.Phase3 = bool.Parse(boolcheck4);
                                NewQbl.Qinfo.Drytime = bool.Parse(boolcheck5);
                                NewQbl.Qinfo.JointColour = bool.Parse(boolcheck6);
                                NewQbl.Qinfo.Colour = bool.Parse(boolcheck7);
                                NewQbl.Qinfo.SlipRes = bool.Parse(boolcheck8);
                                NewQbl.Qinfo.QuoteStages = bool.Parse(boolcheck9);
                                NewQbl.Qinfo.ArrowHead = bool.Parse(boolcheck10);
                                NewQbl.Qinfo.AllowCtrlJoins = bool.Parse(boolcheck11);

                                NewQbl.Qinfo.JointCstr = parts[41];
                                NewQbl.Qinfo.Colstr = parts[42];
                                NewQbl.Qinfo.SlipRstr = parts[43];
                                NewQbl.Qinfo.QuoteStagesstr = parts[44];
                                NewQbl.Qinfo.Arrow = parts[45];
                                NewQbl.Qinfo.JointCstr = parts[46];
                            }
                        }
                        if (parts[0] == "heading1")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc1.QcHeading = parts[1];
                            //qc1.QcSubHeading = parts[1];
                        }
                        if (parts[0] == "text1")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc1.QcTextList.Add(parts[1]);
                        }
                        if (parts[0] == "heading2")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc2.QcHeading = parts[1];
                            //qc2.QcSubHeading = parts[1];
                        }
                        if (parts[0] == "text2")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc2.QcTextList.Add(parts[1]);
                        }
                        if (parts[0] == "heading3")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc3.QcHeading = parts[1];
                            //qc3.QcSubHeading = parts[1];
                        }
                        if (parts[0] == "text3")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            qc3.QcTextList.Add(parts[1]);
                        }
                        if (parts[0] == "Gcom")
                        {
                            parts[1] = parts[1].Replace("^", ",");
                            NewQbl.Qinfo.CommentList.Add(parts[1]);
                        }
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (parts[0] == "text1" && Nextparts[0] == "heading1")
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text1" && Nextparts[0] == "heading2")
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text1" && Nextparts[0] == "heading3")
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text1" && Nextparts[0] == "Gcom")
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text1" && (i + 1) == lines.Length)
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }

                            if (parts[0] == "text2" && Nextparts[0] == "heading2")
                            {
                                if (qc2.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q2CommentList.Add(qc2);                                    
                                    qc2 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text2" && Nextparts[0] == "heading3")
                            {
                                if (qc2.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q2CommentList.Add(qc2);
                                    qc2 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text2" && Nextparts[0] == "Gcom")
                            {
                                if (qc2.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q2CommentList.Add(qc2);
                                    qc2 = new QuoteComment();
                                }
                            }

                            if (parts[0] == "text3" && Nextparts[0] == "heading3")
                            {
                                if (qc3.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q3CommentList.Add(qc3);                                   
                                    qc3 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text3" && Nextparts[0] == "Gcom")
                            {
                                if (qc3.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q3CommentList.Add(qc3); 
                                    qc3 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text1" && i + 1 == lines.Length)
                            {
                                if (qc1.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q1CommentList.Add(qc1);
                                    qc1 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text2" && i + 1 == lines.Length)
                            {
                                if (qc2.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q2CommentList.Add(qc2);
                                    qc2 = new QuoteComment();
                                }
                            }
                            if (parts[0] == "text3" && i + 1 == lines.Length)
                            {
                                if (qc3.QcTextList.Count != 0)
                                {
                                    NewQbl.Qinfo.Q3CommentList.Add(qc3);
                                    qc3 = new QuoteComment();
                                }
                            }
                            //if (parts[0] == "text1" && Nextparts[0] == "heading3" && Nextparts[0] == "Gcom" || i + 1 == lines.Length)
                            //{
                            //    if (qc3.QcTextList.Count != 0)
                            //    {
                            //        Add2QuoteList.Add(qc1);
                            //        qc3 = new QuoteComment();
                            //    }
                            //}
                        }
                    }
                }
                Qblimport = NewQbl;
            }            
        }

        //public string ReplaceChar(string Source, string Find, string Replace)
        //{
        //    bool ok = false;
        //    string result = Source;
        //    while (!ok)
        //    {
        //        if (Source.Contains(Find))
        //        {
        //            int Place = Source.IndexOf(Find);
        //            result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
        //            Source = result;
        //        }
        //        else
        //        {
        //            ok = true;
        //            break;
        //        }
        //    }
        //    return result;
        //}
    }
    [Serializable]
    public class ProductSystem
    {
        public string BaseMat { get; set; }
        public string Type { get; set; }
        public string Flag { get; set; } 
        public List<Pack> PackList = new List<Pack>();
    }

    [Serializable]
    public class Pack
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public double Weight { get; set; }
        public double Volume { get; set; }
        public double Coverage { get; set; }
        public string Unit { get; set; }
        public double Meterage { get; set; }
        public int Packid { get; set; }

        public List<Priceitem> PriceItemList = new List<Priceitem>();
        public FloorOptions Options = new FloorOptions();
        public void AddChild(Priceitem PIin)
        {
            PriceItemList.Add(PIin);
        }
        public string ParentSystemName;
    }

    public class Priceitem
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public double Weight { get; set; }
        public double Volume { get; set; }
        public double Coverage { get; set; }
        public string Unit { get; set; }
        public double Meterage { get; set; }
        public string productDate { get; set; }
    }

    public class HourRates
    {
        public double Grind1 { get; set; }
        public double Grind2 { get; set; }
        public double RollCprep1 { get; set; }
        public double RollCprep2 { get; set; }
        public double RollCprep3 { get; set; }
        public double Trowel4mm { get; set; }
        public double Trowel6mm { get; set; }
        public double coving75mm { get; set; }
        public double coving50mm { get; set; }
        public double FlakeCove { get; set; }
        public double Chevron { get; set; }
        public double Linemark { get; set; }
        public double Subfill { get; set; }
        public double Polyvac { get; set; }
        public double HrMulti { get; set; }
        public double CostMarkUp { get; set; }
        public double PcGnc { get; set; }
        public double pcMp { get; set; }
        public double PcEdgeGnc { get; set; }
        public double pcEdgeMp { get; set; }
        public double GrnSealLightpr { get; set; }
        public double GrnSealMedpr { get; set; }
        public double GrnSealFullpr { get; set; }
        public double MpSealLightpr { get; set; }
        public double MpSealMedpr { get; set; }
        public double MpSealFullpr { get; set; }

    }

    public class Company
    {
        public string CompName { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string Suburb { get; set; }
        public string State { get; set; }
        public  string PostCode { get; set; }
        public string ContactName { get; set; }
        public string ContactNum { get; set; }
    }
    public class Costing
    {
        public double packsAmnt { get; set; }
        public double Ammount { get; set; }
        public string symbol { get; set; }
        public string inputname { get; set; }
        public string PackSystemName { get; set; }
        public double cost { get; set; }
        public double price { get; set; }
        public double Packvol { get; set; }
    }


    public class CSstage
    {
        public string Staging { get; set; }
        public string Systems { get; set; }
        public double Area { get; set; }
        public double Lm { get; set; }
        public double packsAmnt { get; set; }
        public double Ammount { get; set; }
        public string symbol { get; set; }
        public string inputname { get; set; }
        public string PackSystemName { get; set; }
        public double cost { get; set; }
        public double price { get; set; }
        public double Packvol { get; set; }
    }

    public class stagedhrs
    {
        public string Staging { get; set; }
        public string Systems { get; set; }
        public double Area { get; set; }
        public double Lm { get; set; }
        public double packsAmnt { get; set; }
        public string inputname { get; set; }
        public string PackSystemName { get; set; }

    }
    public class Hrrates
    {
        public string hrsname { get; set; }
        public double hrsWeight { get; set; }
    }

    public class HrCombo
    {
        public string SysName { get; set; }
        public string Method { get; set; }
        public string Stage { get; set; }
        public double Area { get; set; }
        public double Length { get; set; }
        public double Rate { get; set; }
        public double Result { get; set; }
    }

    public class QuoteFS
    {
        public string Name { get; set; }
        public string savedHrRate { get; set; }
        public string savedHradd { get; set; }
        public string savedMarkup { get; set; }
        public string savedAdds { get; set; }
        public string SaveDate { get; set; }
        
    }
    public class EstHoursSave
    {
        public string stageName { get; set; }
        public string systemName { get; set; }
        public string m2 { get; set; }
        public string lm { get; set; }
        public string method { get; set; }
        public string techHours { get; set; }
        public string Hours { get; set; }

    }
    public class Quote_Block
    {
        public QuoteInfo Qinfo = new QuoteInfo();
        public List<CSstage> qbStagelist = new List<CSstage>();
        public List<EstHoursSave> estHrslist = new List<EstHoursSave>();
        public List<Adds> qbAdditionslist = new List<Adds>();
        public string QName { get; set; }
        public string savedHrRate { get; set; }
        public string savedHradd { get; set; }
        public string savedMarkup { get; set; }
        public string savedAdds { get; set; }
        public string SaveDate { get; set; }
        public void SaveQuotestageList(string filepath)
        {
            string FileName = QName + "-QuoteStageData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "" +
                "Qname,Name,Ammount,Area,Lm,Cost,PackAmnt,PackSysName,Stage,Systems,Price";
            List<string> outList = new List<string>();

            foreach (CSstage csItem in qbStagelist)
            {
                outList.Add(QName + "," + csItem.inputname + "," + csItem.Ammount + "," + csItem.Area + "," + csItem.Lm + "," +
                                csItem.cost + "," + csItem.packsAmnt + "," + csItem.PackSystemName + "," + csItem.Packvol + "," + csItem.Staging + "," +
                                csItem.symbol + "," + csItem.Systems + "," + csItem.price);
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadQuotestageList(string filepath)
        {           
            string FileName = QName + "-QuoteStageData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(','); 
                    CSstage CSitem = new CSstage();
                    CSitem.inputname = parts[1];
                    CSitem.Ammount = double.Parse(parts[2]);
                    CSitem.Area = double.Parse(parts[3]);
                    CSitem.Lm = double.Parse(parts[4]);
                    CSitem.cost = double.Parse(parts[5]);
                    CSitem.packsAmnt = double.Parse(parts[6]);
                    CSitem.PackSystemName = parts[7];
                    CSitem.Packvol = double.Parse(parts[8]);
                    CSitem.Staging = parts[9];
                    CSitem.symbol = parts[10];
                    CSitem.Systems = parts[11];
                    CSitem.price = double.Parse(parts[12]);
                    qbStagelist.Add(CSitem);
                }
            }
        }
        public void SaveQuoteHrsList(string filepath)
        {
            string FileName = QName + "-QuoteHrsData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "" +
                "Qname,Hours,Lm,m2,Method,StageName,systemName,Techhrs";
            List<string> outList = new List<string>();

            foreach (EstHoursSave hrItem in estHrslist)
            {
                outList.Add(QName + "," + hrItem.Hours + "," + hrItem.lm + "," + hrItem.m2 + "," + hrItem.method + "," +
                                hrItem.stageName + "," + hrItem.systemName + "," + hrItem.techHours);
            }

            //EstHoursSave HRitem = new EstHoursSave();
            //HRitem.Hours = qPackname;
            //HRitem.lm = qAmmount.ToString();
            //HRitem.m2 = qArea.ToString();
            //HRitem.method = qLm.ToString();
            //HRitem.stageName = qCost.ToString();
            //HRitem.systemName = qPackAmt.ToString();
            //HRitem.techHours = qPackSystemName;
            //Qbl.estHrslist.Add(HRitem);

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadQuoteHrsList(string filepath)
        {
            Pack nP = new Pack();
            string FileName = QName + "-QuoteHrsData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');

                    EstHoursSave HRitem = new EstHoursSave();
                    HRitem.Hours = parts[1];
                    HRitem.lm = parts[2];
                    HRitem.m2 = parts[3];
                    HRitem.method = parts[4];
                    HRitem.stageName = parts[5];
                    HRitem.systemName = parts[6];
                    HRitem.techHours = parts[7];
                    estHrslist.Add(HRitem);
                }
            }
        }
    }


    public class QuoteBlock
    {
        public string QName { get; set; }
       // public string Name { get; set; }
        public string savedHrRate { get; set; }
        public string savedHradd { get; set; }
        public string savedMarkup { get; set; }
        public string savedAdds { get; set; }
        public string SaveDate { get; set; }
        public List<Pack> QFPackList = new List<Pack>();
        public List<Pack> QCPackList = new List<Pack>();
        public List<Pack> QSFPackList = new List<Pack>();
        public List<Pack> QLMPackList = new List<Pack>();
        public List<Pack> QCVPackList = new List<Pack>();
        public List<Pack> QPCPackList = new List<Pack>();
        public QuoteInfo Qinfo = new QuoteInfo();

        public void SaveFPackData(string filepath)
        {
            string FileName = QName + "-FPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);

            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QFPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadFPackData(string filepath)
        {

            string FileName = QName + "-FPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);

            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                Pack nP = new Pack();
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddFPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddFPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        Pi.productDate = parts[8];
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddFPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddFPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddFPackToList(Pack newPack)
        {
            //bool itemInList = false;
            //foreach (Pack pItem in QFPackList)
            //{
            //    if (newPack.Name == pItem.Name)
            //    {
            //        itemInList = true;
            //    }
            //}
            //if (!itemInList)
            //{
            QFPackList.Add(newPack);
            //}
        }
        public void SaveCPackData(string filepath)
        {

            string FileName = QName + "-CPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QCPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadCPackData(string filepath)
        {
            string FileName = QName + "-CPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                Pack nP = new Pack();

                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddCPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddCPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        Pi.productDate = parts[8];
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddCPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddCPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddCPackToList(Pack newPack)
        {
            bool itemInList = false;
            foreach (Pack pItem in QCPackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                QCPackList.Add(newPack);
            }
        }
        public void SaveSFPackData(string filepath)
        {
            string FileName = QName + "-SFPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QSFPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," + prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadSFPackData(string filepath)
        {
            Pack nP = new Pack();
            string FileName = QName + "-SFPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddSFPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddSFPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        Pi.productDate = parts[8];
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddSFPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddSFPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddSFPackToList(Pack newPack)
        {
            bool itemInList = false;
            foreach (Pack pItem in QCPackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                QSFPackList.Add(newPack);
            }
        }
        public void SaveLMPackData(string filepath)
        {
            string FileName = QName + "-LMPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QLMPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadLMPackData(string filepath)
        {
            Pack nP = new Pack();
            string FileName = QName + "-LMPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddLMPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddLMPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        Pi.productDate = parts[8];
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddLMPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddLMPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddLMPackToList(Pack newPack)
        {
            bool itemInList = false;
            foreach (Pack pItem in QLMPackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                QLMPackList.Add(newPack);
            }
        }
        public void SaveCVPackData(string filepath)
        {
            string FileName = QName + "-CVPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QCVPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadCVPackData(string filepath)
        {
            Pack nP = new Pack();
            string FileName = QName + "-CVPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddLMPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddCVPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        Pi.productDate = parts[8];
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddCVPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddCVPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddCVPackToList(Pack newPack)
        {
            bool itemInList = false;
            foreach (Pack pItem in QCVPackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                QCVPackList.Add(newPack);
            }
        }
        public void SavePCPackData(string filepath)
        {
            string FileName = QName + "-PCPackData.csv";            
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            System.IO.File.WriteAllBytes(@pathString2, new byte[0]);
            string titleStr = "Type,Name,Price,Weight,Volume,Meterage,Coverage,Unit,Grind,Rollcoat,Trowel,Polivac,CovingType,ProductDate";
            List<string> outList = new List<string>();
            foreach (Pack pItem in QPCPackList)
            {
                outList.Add("Pack" + "," + pItem.Name + "," + pItem.Price + "," + pItem.Weight + "," +
                                pItem.Volume + "," + pItem.Meterage + "," + pItem.Coverage + "," + pItem.Unit + "," + pItem.Options.Grind + "," +
                                pItem.Options.RollCoat + "," + pItem.Options.Trowel + "," + pItem.Options.Polivac + "," + pItem.Options.CovingType);
                foreach (Priceitem prItem in pItem.PriceItemList)
                {
                    outList.Add("sub" + "," + prItem.Name + "," + prItem.Price + "," + prItem.Weight + "," +
                                 prItem.Volume + "," + pItem.Meterage + "," + prItem.Coverage + "," + prItem.Unit + "," + prItem.productDate);
                }
            }

            string[] outArray = new string[outList.Count() + 1];
            outArray[0] = titleStr;
            int j = 1;
            foreach (string stritem in outList)
            {
                outArray[j] = stritem;
                j++;
            }
            System.IO.File.WriteAllLines(@pathString2, outArray);
        }
        public void LoadPCPackData(string filepath)
        {
            Pack nP = new Pack();
            string FileName = QName + "-PCPackData.csv";
            string PTH2 = "Quotes";
            string path = System.IO.Path.Combine(filepath, PTH2);
            string pathString2 = System.IO.Path.Combine(path, FileName);
            if (!File.Exists(@pathString2))
            {
                System.IO.File.Create(@pathString2);
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(@pathString2);
                for (int i = 1; i < lines.Length; i++)
                {
                    var tmpValue = lines[i];
                    var parts = tmpValue.Split(',');
                    if (parts[0] == "Pack")
                    {
                        nP.Name = parts[1];
                        nP.Price = double.Parse(parts[2]);
                        nP.Weight = double.Parse(parts[3]);
                        nP.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        nP.Coverage = double.Parse(parts[6]);
                        nP.Unit = parts[7];
                        nP.Options.Grind = parts[8];
                        nP.Options.RollCoat = parts[9];
                        nP.Options.Trowel = parts[10];
                        nP.Options.Polivac = int.Parse(parts[11]);
                        nP.Options.CovingType = parts[12];
                        //nP.Meterage = 0;

                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddLMPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddPCPackToList(nP);
                        }
                    }
                    if (parts[0] == "sub")
                    {
                        Priceitem Pi = new Priceitem();
                        Pi.Name = parts[1];
                        Pi.Price = double.Parse(parts[2]);
                        Pi.Weight = double.Parse(parts[3]);
                        Pi.Volume = double.Parse(parts[4]);
                        nP.Meterage = double.Parse(parts[5]);
                        Pi.Coverage = double.Parse(parts[6]);
                        Pi.Unit = parts[7];
                        Pi.Meterage = 0;
                        nP.AddChild(Pi);
                        if (i + 1 < lines.Length)
                        {
                            var NexttmpValue = lines[i + 1];
                            var Nextparts = NexttmpValue.Split(',');
                            if (Nextparts[0] == "Pack")
                            {
                                AddPCPackToList(nP);
                                nP = new Pack();
                            }
                        }
                        else if (i + 1 == lines.Length)
                        {
                            AddPCPackToList(nP);
                        }
                    }
                }
            }
        }
        public void AddPCPackToList(Pack newPack)
        {
            bool itemInList = false;
            foreach (Pack pItem in QPCPackList)
            {
                if (newPack.Name == pItem.Name)
                {
                    itemInList = true;
                }
            }
            if (!itemInList)
            {
                QPCPackList.Add(newPack);
            }
        }

    }

    public class Quotingdata
    {
        public DateTime dt;
        public float QuoteAmnt;
    }

    public class QuoteInfo
    {
        public string QuoteName { get; set; }
        public string Jobnum { get; set; }
        public string JobDate { get; set; }
        public string CompanyName { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string Suburb { get; set; }
        public string State { get; set; }
        public string Postcode { get; set; }
        public string ClientName { get; set; }
        public string ClientNum{ get; set; }
        public string Project { get; set; }
        public string TotalEst { get; set; }
        public string Q1Allowance { get; set; }
        public string Q1Lm { get; set; }
        public string Q1Days { get; set; }
        public string Q1Heading { get; set; }
        public string Q1SlipRate { get; set; }
        public string Q1StageVal { get; set; }
        public string Q1Warrenty { get; set; }
        public List<QuoteComment> Q1CommentList = new List<QuoteComment>();
        public bool Q2Enable { get; set; }
        public string Q2Allowance { get; set; }
        public string Q2Lm { get; set; }
        public string Q2Days { get; set; }
        public string Q2Heading { get; set; }
        public string Q2SlipRate { get; set; }
        public string Q2StageVal { get; set; }
        public string Q2Warrenty { get; set; }
        public List<QuoteComment> Q2CommentList = new List<QuoteComment>();
        public bool Q3Enable { get; set; }
        public string Q3Allowance { get; set; }
        public string Q3Lm { get; set; }
        public string Q3Days { get; set; }
        public string Q3Heading { get; set; }
        public string Q3SlipRate { get; set; }
        public string Q3StageVal { get; set; }
        public string Q3Warrenty { get; set; }
        public List<QuoteComment> Q3CommentList = new List<QuoteComment>();
        public bool tnc { get; set; }
        public string Warrenty { get; set; }
        public string CompletedOver{ get; set; }

        // General comments
        public List<String> CommentList = new List<String>();
        public bool Rubbish { get; set; }
        public bool Phase3 { get; set; }
        public bool Drytime { get; set; }
        public bool JointColour { get; set; }
        public string JointCstr { get; set; }
        public bool Colour { get; set; }
        public string Colstr { get; set; }
        public bool SlipRes { get; set; }
        public string SlipRstr { get; set; }
        public bool QuoteStages { get; set; }
        public string QuoteStagesstr { get; set; }
        public bool ArrowHead { get; set; }
        public string Arrow{ get; set; }
        public bool AllowCtrlJoins { get; set; }
        public string AllowCtrlstr { get; set; }

    }

    public class FloorOptions
    {
        public string Grind { get; set; }
        public string RollCoat { get; set; }
        public string Trowel { get; set; }
        public int Polivac { get; set; }
        public string CovingType { get; set; }
        //public string Coving { get; set; }
    }
        
    public class datapack
    {
        //public int abc { get; set; }
        public List<ProductSystem> SystemData = new List<ProductSystem>();
        public List<Pack> PackData = new List<Pack>();
        public List<Priceitem> PriceItemData = new List<Priceitem>();
    }
    public class QuoteComment
    {
        public string QcHeading { get; set; }
        public string QcSubHeading { get; set; }
        public List<string> QcTextList = new List<string>();
    }

    public class QcContainer
    {
        public string QcHeading { get; set; }
        public List<QuoteComment> QList = new List<QuoteComment>();
    }
    public class Adds
    {
        public string ItemName { get; set; }
        public string Desc { get; set; }
        public string Ammount { get; set; }
        public string Cost { get; set; }

    }

}
