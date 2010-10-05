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
    public partial class CreateOpportunity : Form
    {
        SDataPayload contactPayload;
        public CreateOpportunity(SDataPayload payload)
        {
            InitializeComponent();
            contactPayload = payload;
        }

        private void CreateOpportunity_Load(object sender, EventArgs e)
        {

        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                string ownerId = UserNameToOwnerId.GetId(service.UserName);
                string userId = UserNameToId.GetId(service.UserName);

                var entry = new AtomEntry();
                var payload = new SDataPayload
                {
                    ResourceName = "Opportunity",
                    Namespace = "http://schemas.sage.com/dynamic/2007",
                    Values = {
                                {"Description", txtDescription.Text},
                                {"Account", (SDataPayload)contactPayload.Values["Account"]},
                                {"Owner", new SDataPayload{ Key = ownerId, ResourceName="Owner"}},
                                {"AccountManager", new SDataPayload{ Key = userId, ResourceName="AccountManager"}},
                                {"Contacts", 
                                    new SDataPayloadCollection {
                                        new SDataPayload {
                                            ResourceName = "OpportunityContact",
                                            Values = {{"Contact",new SDataPayload{ Key = contactPayload.Key}},
                                                      {"IsPrimary","true"}}
                                        }
                                    }
                                }
                             }
                };

                entry.SetSDataPayload(payload);
                var request = new SDataSingleResourceRequest(service, entry) { ResourceKind = "Opportunities" };
                AtomEntry result = request.Create();

                this.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
