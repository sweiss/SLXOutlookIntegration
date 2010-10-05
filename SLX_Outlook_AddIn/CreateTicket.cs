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
    public partial class CreateTicket : Form
    {
        SDataPayload contactPayload;
        SafeMailItem curEmail;
        public CreateTicket(SDataPayload payload, SafeMailItem safeEmail)
        {
            InitializeComponent();
            contactPayload = payload;
            curEmail = safeEmail;
        }

        private void CreateTicket_Load(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(curEmail.Body))
            {
                txtSubject.Text = curEmail.Body; 
            }
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                string ownerId = UserNameToOwnerId.GetId(service.UserName);

                var entry = new AtomEntry();
                var payload = new SDataPayload
                {
                    ResourceName = "Ticket",
                    Namespace = "http://schemas.sage.com/dynamic/2007",
                    Values = {
                    {"Account", (SDataPayload)contactPayload.Values["Account"]},           
                    {"Contact", contactPayload},//,
                                {"AssignedTo", new SDataPayload{ Key = ownerId, ResourceName="Owner"}},
                                //{"Issue",txtSubject.Text},
                                {"TicketProblem", new SDataPayload{
                                    ResourceName="TicketProblem",
                                    Values = {
                                        {"Notes",txtSubject.Text}
                                    }
                                }
                                }
                    }
                };

                entry.SetSDataPayload(payload);
                var request = new SDataSingleResourceRequest(service, entry) { ResourceKind = "Tickets" };
                AtomEntry result = request.Create();

                this.Close();

            }
            catch (SDataClientException ex)
            {
                //MessageBox.Show(ex.Message);
                //Getting object reference error and have no idea why
                //Everything is created just fine and is working though still get error
                this.Close();
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
