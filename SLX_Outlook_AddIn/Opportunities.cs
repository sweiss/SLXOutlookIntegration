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
using Sage.SData.Client.Extensions;
using System.Diagnostics;

namespace SLX_Outlook_AddIn
{
    public partial class Opportunities : Form
    {
        string id;
        
        public Opportunities(string contactId)
        {
            InitializeComponent();
            id = contactId;
            
        }

        private void Opportunities_Load(object sender, EventArgs e)
        {
            LoadOpportunities();
        }

        private void LoadOpportunities()
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                SDataResourceCollectionRequest oppoContactsCollection = new SDataResourceCollectionRequest(service);

                oppoContactsCollection.ResourceKind = "opportunitycontacts";
                oppoContactsCollection.QueryValues.Add("where", "Contact.Id eq '" + id + "'");

                AtomFeed oppoContactsFeed = oppoContactsCollection.Read();

                if (oppoContactsFeed.Entries.Count() > 0)
                {
                    
                    DataTable table = new DataTable();
                    
                    table.Columns.Add("Id");
                    table.Columns.Add("Description");

                    
                    foreach (AtomEntry entry in oppoContactsFeed.Entries)
                    {
                        SDataPayload oppoContact = entry.GetSDataPayload();

                        SDataPayload tempOppo = (SDataPayload)oppoContact.Values["Opportunity"];
                        
                        SDataSingleResourceRequest oppo = new SDataSingleResourceRequest(service);

                        oppo.ResourceKind = "Opportunities";

                        oppo.ResourceSelector = "('" + tempOppo.Key + "')";

                        AtomEntry oppoEnty = oppo.Read();

                        SDataPayload opportunitiy = (SDataPayload)oppoEnty.GetSDataPayload();

                        DataRow dr = table.NewRow();

                        dr[0] = opportunitiy.Key.ToString();
                        dr[1] = opportunitiy.Values["Description"].ToString();

                        table.Rows.Add(dr);
                    }

                    grdOpportunities.DataSource = table;
                    grdOpportunities.Columns[2].Visible = false;
                    grdOpportunities.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void grdOpportunities_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)//Details
            {
                DisplayDetails(e);
            }

            if (e.ColumnIndex == 1)//Open
            {
                OpenOpportunity(e);
            }
        }

        private void DisplayDetails(DataGridViewCellEventArgs e)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                SDataSingleResourceRequest oppo = new SDataSingleResourceRequest(service);

                oppo.ResourceKind = "Opportunities";

                oppo.ResourceSelector = "('" + grdOpportunities.Rows[e.RowIndex].Cells[2].Value.ToString() + "')";

                AtomEntry oppoEnty = oppo.Read();

                SDataPayload opportunitiy = (SDataPayload)oppoEnty.GetSDataPayload();

                txtOppoDesc.Text = opportunitiy.Values["Description"].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OpenOpportunity(DataGridViewCellEventArgs e)
        {
            try
            {
                Process iexplore = new Process();

                string tempURL = Properties.Settings.Default.SDATA;

                tempURL += "/{0}/Opportunity.aspx?entityid={1}";

                iexplore.StartInfo.FileName = "iexplore.exe";
                iexplore.StartInfo.Arguments = String.Format(tempURL, "SlxClient", grdOpportunities.Rows[e.RowIndex].Cells[2].Value.ToString());
                iexplore.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

}

