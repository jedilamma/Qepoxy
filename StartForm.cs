using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SuperSimpleTcp;

namespace Qepoxy
{
    public partial class StartForm : Form
    {
        
        Backend newBackend = new Backend();
        SimpleTcpClient client;
        Main QepoxyProgram;
        string ExternalIp;
        public StartForm()
        {
            InitializeComponent();
            getExternalIp();
            Usernametb.Text = "";
            Passwordtb.PasswordChar = '●';
            Passwordtb.Text = "";

            Usernametb.Text = "Admin";
            Passwordtb.Text = "Pass";

            //newBackend.lists.attemptConnect();
            ServerConnectCb.Text = "192.168.56.1";
            //ServerConnectCb.Text = "qepoxyserver.ddns.net"; 
            ServerConnectCb.Items.Add(newBackend.lists.server);
            ServerConnectCb.Items.Add("qepoxyserver.ddns.net");
            ServerConnectCb.Items.Add("192.168.0.6");
            ServerConnectCb.Items.Add("192.168.0.2");
            ServerConnectCb.Items.Add("192.168.56.1");
            InfiniteLoop();
        }
        public void OpenRoutine()
        {
            Loginbtn.Enabled = false;
            QepoxyProgram = new Main(newBackend);
            QepoxyProgram.Owner = this;
            QepoxyProgram.Show();
            Usernametb.Enabled = false;
            Passwordtb.Enabled = false;
            Loginbtn.Enabled = false;            

        }
        public async void CloseRoutine()
        {
            this.Show();
            await Task.Delay(1000);           
            Close();
        }

        private void StartProgram()
        {
            if (newBackend.Startup)
            {
                OpenRoutine();
                newBackend.Startup = false;
            }
        }

       
    private void Loginbtn_Click(object sender, EventArgs e)
        {
            if (client.IsConnected)
            {
                client.Send("QID" + "," + Usernametb.Text + "," + Passwordtb.Text + "," + ExternalIp);
            }
        }

        private void Connect1_Click(object sender, EventArgs e)
        {
            if(ServerConnectCb.Text != "")
            {
                ConnectBtn.Enabled = false;
                client = new(ServerConnectCb.Text, 8500);
                client.Events.DataReceived += Events_DataRecieved;
                client.Events.Connected += Events_Connected;
                client.Events.Disconnected += Events_Disconnected;
                AttemptConnect();

                if (client != null)
                {
                    if (client.IsConnected)
                    {
                        //ConnectBtn.Enabled = false;
                    }
                    else
                    {
                        try
                        {
                            client.Connect();
                            client.Send("CompName = "+Environment.MachineName.ToString()); ;
                            //if (client.IsConnected)
                            //{
                            //    client.Send("QID" + "," + Usernametb.Text + "," + Passwordtb.Text + "," + ExternalIp);
                            //}
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Message,", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            pictureBox2.Image = Qepoxy.Properties.Resources.red1;
                            ConnectBtn.Enabled = true;
                        }

                        if (client.IsConnected)
                        {
                            pictureBox2.Image = Qepoxy.Properties.Resources.green1;
                            //Loginbtn.Enabled = true;
                        }
                    }
                }
            }            
        }

        //private void StartForm_Load(object sender, EventArgs e)
        //{
        //    client = new(ServerConnectCb.Text,8500);
        //    client.Events.DataReceived += Events_DataRecieved;
        //    client.Events.Connected += Events_Connected;
        //    client.Events.Disconnected += Events_Disconnected;
        //    AttemptConnect();
        //}
        private async void InfiniteLoop()
        {
            while(true)
            { 
                await Task.Delay(1000);
                if(newBackend.GobalDC && client.IsConnected)
                {
                    pictureBox2.Image = Qepoxy.Properties.Resources.red1;
                    Loginbtn.Enabled = false;
                    ConnectBtn.Enabled = true;
                    DcBtn.Enabled = false;
                    if(QepoxyProgram != null)
                    {
                        QepoxyProgram.ForceClose();
                    }                    
                    client.Disconnect();
                    newBackend.GobalDC = false;
                    newBackend.Startup = true;
                    Usernametb.Enabled = true;
                    Passwordtb.Enabled = true;
                }
            }
        }

        private async void AttemptConnect()
        {
            await Task.Delay(1000);
            try
            {
                client.Connect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message,", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
           
            if (client.IsConnected)
            {
                pictureBox2.Image = Qepoxy.Properties.Resources.green1;
                //Loginbtn.Enabled = true;
                ConnectBtn.Enabled = false;
                DcBtn.Enabled = true;
            }
            else
            {
                ConnectBtn.Enabled = true;
                DcBtn.Enabled = false;
                pictureBox2.Image = Qepoxy.Properties.Resources.red1;
            }
        }

        private void Events_DataRecieved(object sender, DataReceivedEventArgs e)
        {
            // process data right here
            string indata = Encoding.UTF8.GetString(e.Data);
            if (indata == "PWaccepted")
            {
                Invoke((MethodInvoker)delegate { StartProgram(); });               

            }
            if (indata == "Connected")
            {
                this.Invoke((MethodInvoker)delegate { Loginbtn.Enabled = true; });
                this.Invoke((MethodInvoker)delegate {
                    Usernametb.Enabled = true;
                });
                this.Invoke((MethodInvoker)delegate {
                        Passwordtb.Enabled = true;
                });



            }
            if (indata == "LoginFail")
            {
                this.Invoke((MethodInvoker)delegate { Statuslbl.Text = "Login Failed Bad Username or Password"; });
                               
            }
            if (indata == "AdminDisconnect")
            {
                this.Invoke((MethodInvoker)delegate { newBackend.GobalDC = true; });
            }
        }

        private void Events_Connected(object sender, ConnectionEventArgs e)
        {
            pictureBox2.Image = Qepoxy.Properties.Resources.green1;
            //Loginbtn.Enabled = true;
        }
        private void Events_Disconnected(object sender, ConnectionEventArgs e)
        {
            pictureBox2.Image = Qepoxy.Properties.Resources.red1;
            //Loginbtn.Enabled = false;
        }

        private void DcBtn_Click(object sender, EventArgs e)
        {
            if(client!= null)
            if(client.IsConnected)
            {
                client.Disconnect();
                pictureBox2.Image = Qepoxy.Properties.Resources.red1;
                Loginbtn.Enabled = false;
                ConnectBtn.Enabled = true;
                DcBtn.Enabled = false;
                newBackend.Startup = true;
                Usernametb.Enabled = true;
                Passwordtb.Enabled = true;
            }          
        }

        private void getExternalIp()
        {
            try
            {
                string url = "http://checkip.dyndns.org";
                System.Net.WebRequest req = System.Net.WebRequest.Create(url);
                System.Net.WebResponse resp = req.GetResponse();
                System.IO.StreamReader sr = new System.IO.StreamReader(resp.GetResponseStream());
                string response = sr.ReadToEnd().Trim();
                string[] ipAddressWithText = response.Split(':');
                string ipAddressWithHTMLEnd = ipAddressWithText[1].Substring(1);
                string[] ipAddress = ipAddressWithHTMLEnd.Split('<');
                ExternalIp = ipAddress[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message,", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }          
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            Passwordtb.PasswordChar = '\0';
        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            Passwordtb.PasswordChar = '●';
        }

        private void ServerConnectCb_SelectedIndexChanged(object sender, EventArgs e)
        {
            //client = new(ServerConnectCb.Text, 8500);
            //client.Events.DataReceived += Events_DataRecieved;
            //client.Events.Connected += Events_Connected;
            //client.Events.Disconnected += Events_Disconnected;
            //AttemptConnect();
        }
    }
}
