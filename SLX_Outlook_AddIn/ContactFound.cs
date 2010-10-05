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
using Redemption;

namespace SLX_Outlook_AddIn
{
    public partial class ContactFound : Form
    {
        string contactId;
        AtomFeed contactFeed;
        ISDataService mydataService;
        SDataPayload payload;
        SafeMailItem curEmail;

        public ContactFound(string id, SafeMailItem safeMail)
        {
            InitializeComponent();
            contactId = id;
            curEmail = safeMail;
        }

        private void ContactFound_Load(object sender, EventArgs e)
        {

            try
            {
                if (!String.IsNullOrEmpty(contactId))
                {
                    mydataService = SDataDataService.mydataService();

                    SDataSingleResourceRequest mydataSingleRequest;
                    mydataSingleRequest = new SDataSingleResourceRequest(mydataService);
                    mydataSingleRequest.ResourceKind = "Contacts";
                    mydataSingleRequest.ResourceSelector = "('" + contactId + "')";
                    AtomEntry myContact = mydataSingleRequest.Read();
                    mydataSingleRequest.Entry = myContact;

                    payload = mydataSingleRequest.Entry.GetSDataPayload();

                    if (payload != null)
                    {
                        txtAccount.Text = payload.Values["AccountName"].ToString().Trim();
                        txtFirstName.Text = payload.Values["FirstName"].ToString().Trim();
                        txtLastName.Text = payload.Values["LastName"].ToString().Trim();
                    } 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmdCreate_Click(object sender, EventArgs e)
        {
            if (radioButtonCreateOppo.Checked)
            {
                CreateOpportunity frmCreateOpportunity = new CreateOpportunity(payload);
                frmCreateOpportunity.FormClosed += new FormClosedEventHandler(frmCreateOpportunity_FormClosed);
                frmCreateOpportunity.ShowDialog();
                
            }
            else if (radioButtonCreateTicket.Checked)
            {
                CreateTicket frmCreateTicket = new CreateTicket(payload, curEmail);
                frmCreateTicket.FormClosed += new FormClosedEventHandler(frmCreateTicket_FormClosed);
                frmCreateTicket.ShowDialog();
            }
            else if (radioButtonCreateActivity.Checked)
            {
                CreateActivity frmActivityTicket = new CreateActivity(payload);
                frmActivityTicket.FormClosed += new FormClosedEventHandler(frmActivityTicket_FormClosed);
                frmActivityTicket.ShowDialog();
            }
        }

        void frmActivityTicket_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }

        void frmCreateTicket_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }

        void frmCreateOpportunity_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }
    }
}
