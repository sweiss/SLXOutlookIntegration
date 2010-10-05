using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Core;
using Sage.SData.Client.Extensions;
using Redemption;

namespace SLX_Outlook_AddIn
{
    public partial class CreateAccountContact : Form
    {      
        public CreateAccountContact(Outlook.MailItem mailItem)
        {
            InitializeComponent();

            try
            {
                if (mailItem != null)
                {
                    SafeMailItem safeMail = new SafeMailItem();

                    safeMail.Item = mailItem;

                    txtEmail.Text = safeMail.Sender.SMTPAddress;

                    if (!String.IsNullOrEmpty(safeMail.SenderName))
                    {
                        string[] nameCommaSplit = safeMail.SenderName.Split(',');
                        string[] nameSpaceSplit = safeMail.SenderName.Split(' ');

                        if (nameCommaSplit.Length > 1)
                        {
                            txtFName.Text = nameCommaSplit[1].Trim();
                            txtLName.Text = nameCommaSplit[0].Trim();
                        }
                        else if (nameSpaceSplit.Length > 1)
                        {
                            txtFName.Text = nameSpaceSplit[0].Trim();
                            txtLName.Text = nameSpaceSplit[1].Trim();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Create_Load(object sender, EventArgs e)
        {

        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                var entry = new AtomEntry();
                var payload = new SDataPayload
                {
                    ResourceName = "Account",
                    Namespace = "http://schemas.sage.com/dynamic/2007",
                    Values = {
                        {"AccountName", txtAccount.Text},
                        {"Contacts", new SDataPayloadCollection {
                                new SDataPayload {
                                    ResourceName = "Contact",
                                    Values = {
                                        {"AccountName",  txtAccount.Text},
                                        {"LastName", txtLName.Text},
                                        {"FirstName", txtFName.Text},
                                        {"Email",txtEmail.Text}
                                    }
                                }
                            }
                        }
                    }
                };

                entry.SetSDataPayload(payload);
                var request = new SDataSingleResourceRequest(service, entry) { ResourceKind = "accounts" };
                AtomEntry result = request.Create();

                if(result != null)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.Yes;
                }

                this.Close();

            }
            catch (SDataClientException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
