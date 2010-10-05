using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Sage.SData.Client.Core;
using Sage.SData.Client.Atom;

namespace SLX_Outlook_AddIn
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
            
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {

        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            string temp = string.Empty;
            string url = string.Empty;
            
            if (temp == "/")
            {
                url += "sdata/slx/dynamic/-/";
            }
            else
            {
                url += "/sdata/slx/dynamic/-/";
            }
            
            Properties.Settings.Default.SDATA = txtSdata.Text;
            Properties.Settings.Default.UserName = txtUserName.Text;
            Properties.Settings.Default.Password = txtPassword.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            txtSdata.Text = Properties.Settings.Default.SDATA;
            txtUserName.Text = Properties.Settings.Default.UserName;
            txtPassword.Text = Properties.Settings.Default.Password;
        }

        private void cmdTest_Click(object sender, EventArgs e)
        {
            try
            {
                string userName = txtUserName.Text;
                string password = txtPassword.Text;
                string url = txtSdata.Text;

                string temp = txtSdata.Text.Substring(txtSdata.Text.Length - 1, 1);

                if (temp == "/")
                {
                    url += "sdata/slx/dynamic/-/";
                }
                else
                {
                    url += "/sdata/slx/dynamic/-/";
                }

                ISDataService service;
                service = new SDataService(url, userName, password);

                SDataResourceCollectionRequest sdataCollection = new SDataResourceCollectionRequest(service);

                sdataCollection.ResourceKind = "Accounts";

                AtomFeed accountFeed = sdataCollection.Read();

                if (accountFeed.Entries.Count() > 0)
                {
                    MessageBox.Show("Test Successful");
                }
            }
            catch (SDataClientException ex)
            {
                MessageBox.Show(ex.InnerException.Message);
            }

        }
    }
}
