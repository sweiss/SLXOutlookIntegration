using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Core;
using Sage.SData.Client.Extensions;

namespace SLX_Outlook_AddIn
{
    public partial class CreateActivity : Form
    {
        SDataPayload contactPayload;
        public CreateActivity(SDataPayload payload)
        {
            InitializeComponent();
            contactPayload = payload;
        }

        private void CreateActivity_Load(object sender, EventArgs e)
        {

        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                string contactId = contactPayload.Key; ;
                string contactName = contactPayload.Values["FirstName"].ToString().Trim() + " " + contactPayload.Values["LastName"].ToString().Trim();

                SDataPayload accountPayload = (SDataPayload)contactPayload.Values["Account"];
                string accountId = accountPayload.Key;
                string accountName = contactPayload.Values["AccountName"].ToString().Trim();

                var entry = new AtomEntry();
                var payload = new SDataPayload
                {
                    ResourceName = "Activity",
                    Namespace = "http://schemas.sage.com/dynamic/2007",
                    Values = {
                    {"AccountId", accountId},
                    //{"AccountName", accountName},
                    {"ContactId", contactId},
                    {"Description", txtRegarding.Text}
                    }
                };

                entry.SetSDataPayload(payload);
                var request = new SDataSingleResourceRequest(service, entry) { ResourceKind = "Activities" };
                AtomEntry result = request.Create();

                if (result != null)
                {
                    //MessageBox.Show("Acctivity created");
                }

                this.Close();

            }
            catch (SDataClientException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
