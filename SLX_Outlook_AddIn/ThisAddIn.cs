using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Core;
using Sage.SData.Client.Extensions;
using Redemption;
using System.Text.RegularExpressions;
using System.Diagnostics;


namespace SLX_Outlook_AddIn
{
    public partial class ThisAddIn
    {
        Office.CommandBarButton cmdButton;
        Office.CommandBarButton cmdGotoContact;
        Office.CommandBarButton cmdOpportunities;
        Office.CommandBarButton cmdSettings;

        ISDataService mydataService;
        SDataResourceCollectionRequest mydataCollection;
        //SDataSingleResourceRequest mydataSingleRequest;
        string contactId;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Custom context menu item event managed
            this.Application.ItemContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(MenuItem_ItemContextMenuDisplay);
        }

        public void MenuItem_ItemContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.Selection Selection)
        {
            try
            {
                
                // Commadbarpopup control to context menu item
                Office.CommandBarPopup CustomItem = (Office.CommandBarPopup)CommandBar.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, "Custom Menu Item", CommandBar.Controls.Count + 1, Type.Missing);
                // Added to separate group in context menu
                CustomItem.BeginGroup = true;
                // Set the tag value for the menu
                CustomItem.Tag = "CustomMenuItem";
                // Caption for the context menu item
                CustomItem.Caption = "Custom SLX Integration";
                // Set it to visible
                CustomItem.Visible = true;

                //Website with all the faceid's http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
                cmdGotoContact = (Office.CommandBarButton)CustomItem.Controls.Add(1, missing, missing, missing, true);
                cmdGotoContact.Caption = "Goto Contact";
                cmdGotoContact.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdGotoContact_Click);
                cmdGotoContact.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                cmdGotoContact.FaceId = 2103;

                //Website with all the faceid's http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
                cmdButton = (Office.CommandBarButton)CustomItem.Controls.Add(1, missing, missing, missing, true);
                cmdButton.Caption = "Create Opportunity / Ticket / Activity";
                cmdButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdButton_Click);
                cmdButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                cmdButton.FaceId = 0577;

                //Website with all the faceid's http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
                cmdOpportunities = (Office.CommandBarButton)CustomItem.Controls.Add(1, missing, missing, missing, true);
                cmdOpportunities.Caption = "Opportunity List";
                cmdOpportunities.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdOpportunities_Click);
                cmdOpportunities.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                cmdOpportunities.FaceId = 0008;

                cmdSettings = (Office.CommandBarButton)CustomItem.Controls.Add(1, missing, missing, missing, true);
                cmdSettings.Caption = "Settings";
                cmdSettings.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdSettings_Click);
                cmdSettings.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                cmdSettings.FaceId = 0212;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void cmdGotoContact_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Application.ActiveExplorer().Selection.Count > 1)
                {
                    MessageBox.Show("Please only select a single email.");
                    return;
                }

                Outlook.MailItem mail = Application.ActiveExplorer().Selection[1] as Outlook.MailItem;

                SafeMailItem safeMail = new SafeMailItem();

                safeMail.Item = mail;

                string add = safeMail.Sender.SMTPAddress;

                if (!String.IsNullOrEmpty(add))
                {
                    Regex regex = new Regex("^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}$", RegexOptions.IgnoreCase);

                    if (regex.IsMatch(add))
                    {
                        contactId = EmailSearch(add);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

                if (contactId == null)
                {
                    MessageBox.Show("Could not find contact");
                }
                else
                {
                    Process iexplore = new Process();

                    string tempURL = Properties.Settings.Default.SDATA;

                    tempURL += "/{0}/Contact.aspx?entityid={1}";

                    iexplore.StartInfo.FileName = "iexplore.exe";
                    iexplore.StartInfo.Arguments = String.Format(tempURL, "SlxClient", contactId);
                    iexplore.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }  
        }

        void cmdOpportunities_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Application.ActiveExplorer().Selection.Count > 1)
                {
                    MessageBox.Show("Please only select a single email.");
                    return;
                }

                Outlook.MailItem mail = Application.ActiveExplorer().Selection[1] as Outlook.MailItem;

                SafeMailItem safeMail = new SafeMailItem();

                safeMail.Item = mail;

                string add = safeMail.Sender.SMTPAddress;

                if (!String.IsNullOrEmpty(add))
                {
                    Regex regex = new Regex("^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}$", RegexOptions.IgnoreCase);

                    if (regex.IsMatch(add))
                    {
                        contactId = EmailSearch(add);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

                if (contactId == null)
                {
                    MessageBox.Show("Could not find contact");
                }
                else
                {
                    Form frmOpportunities = new Opportunities(contactId);
                    frmOpportunities.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }  
        }

        void cmdButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (Application.ActiveExplorer().Selection.Count > 1)
                {
                    MessageBox.Show("Please only select a single email.");
                    return;
                }
                
                Outlook.MailItem mail = Application.ActiveExplorer().Selection[1] as Outlook.MailItem;

                SafeMailItem safeMail = new SafeMailItem();

                safeMail.Item = mail;

                string add = safeMail.Sender.SMTPAddress;

                if (!String.IsNullOrEmpty(add))
                {
                    Regex regex = new Regex("^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}$", RegexOptions.IgnoreCase);

                    if (regex.IsMatch(add))
                    {
                        contactId = EmailSearch(add);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

                if (contactId == null)
                {
                    Form frmCreate = new CreateAccountContact(mail);
                    frmCreate.FormClosed += new FormClosedEventHandler(frmCreate_FormClosed);
                    frmCreate.ShowDialog();
                }
                else
                {
                    Form frmContactFound = new ContactFound(contactId, safeMail);
                    frmContactFound.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void frmCreate_FormClosed(object sender, FormClosedEventArgs e)
        {
            CreateAccountContact frm = (CreateAccountContact)sender;

            Outlook.MailItem mail = Application.ActiveExplorer().Selection[1] as Outlook.MailItem;

            SafeMailItem safeMail = new SafeMailItem();

            safeMail.Item = mail;

            string add = safeMail.Sender.SMTPAddress;

            contactId = EmailSearch(add);

            if ((frm.DialogResult == DialogResult.Yes) & (!String.IsNullOrEmpty(contactId)))
            {
                Form frmContactFound = new ContactFound(contactId, safeMail);
                frmContactFound.ShowDialog();
            }
        }

        void cmdSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Form frmSettings = new Settings();
            frmSettings.ShowDialog();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private string EmailSearch(string address)
        {
            try
            {
                mydataService = SDataDataService.mydataService();

                mydataCollection = new SDataResourceCollectionRequest(mydataService);
                mydataCollection.ResourceKind = "Contacts";
                mydataCollection.QueryValues.Add("where", "Email eq '" + address + "'");
                AtomFeed contactfeed = mydataCollection.Read();
                string contactId;

                if (contactfeed.Entries.Count() > 0)
                {
                    foreach (AtomEntry entry in contactfeed.Entries)
                    {
                        string tempURI = entry.Id.Uri.AbsoluteUri;
                        contactId = tempURI.Substring(tempURI.IndexOf("'") + 1, tempURI.LastIndexOf("'") - tempURI.IndexOf("'") - 1);
                        return contactId;
                    }
                }

                return null;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        private bool testSettings()
        {
            try
            { 

                string userName = Properties.Settings.Default.UserName;
                string password = Properties.Settings.Default.Password;
                string url = Properties.Settings.Default.SDATA;

                if (String.IsNullOrEmpty(userName) || String.IsNullOrEmpty(url))
                {
                    return false;
                }

                string temp = url.Substring(url.Length - 1, 1);

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
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (SDataClientException ex)
            {
                return false;
            }
            }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
