using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sage.SData.Client.Core;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Extensions;

    class CreateOpportunityContact
    {
        public static string Create(string oppoId, string contactId)
        {
            try
            {
                ISDataService service;
                service = SDataDataService.mydataService();

                var entry = new AtomEntry();
                var payload = new SDataPayload
                {
                    ResourceName = "OpportunityContact",
                    Namespace = "http://schemas.sage.com/dynamic/2007",
                    Values = {
                                {"Contact", new SDataPayload{ Key = contactId, ResourceName="Contact"}},
                                {"Opportunity", new SDataPayload{ Key = oppoId, ResourceName="Opportunity"}}
                             }
                };

                entry.SetSDataPayload(payload);
                var request = new SDataSingleResourceRequest(service, entry) { ResourceKind = "OpportunityContacts" };
                AtomEntry result = request.Create();

                if (result != null)
                {
                    return result.Id.ToString();
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }

