using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Core;
using Sage.SData.Client.Extensions;


    class UserNameToOwnerId
    {
        public static string GetId(string userName)
        {
            try
            {
                ISDataService mydataService = SDataDataService.mydataService();
                string ownerId = string.Empty;

                SDataResourceCollectionRequest mydataCollection = new SDataResourceCollectionRequest(mydataService);
                mydataCollection.ResourceKind = "Users";
                mydataCollection.QueryValues.Add("where", "UserName eq '" + userName + "'");
                AtomFeed usersFeed = mydataCollection.Read();
                string userId = string.Empty;

                if (usersFeed.Entries.Count() > 0)
                {
                    foreach (AtomEntry entry in usersFeed.Entries)
                    {
                        SDataPayload payLoad = entry.GetSDataPayload();
                        string tempURI = entry.Id.Uri.AbsoluteUri;
                        userId = tempURI.Substring(tempURI.IndexOf("'") + 1, tempURI.LastIndexOf("'") - tempURI.IndexOf("'") - 1);
                        break;
                    }
                }

                mydataCollection = null;
                mydataCollection = new SDataResourceCollectionRequest(mydataService);
                mydataCollection.ResourceKind = "Owners";
                AtomFeed ownersFeed = mydataCollection.Read();

                if (ownersFeed.Entries.Count() > 0)
                {
                    foreach (AtomEntry entry in ownersFeed.Entries)
                    {
                        SDataPayload payLoad = entry.GetSDataPayload();
                        SDataPayload userPayload = (SDataPayload)payLoad.Values["User"];

                        if (userPayload != null)
                        {
                            if (userPayload.Key == userId)
                            {
                                ownerId = payLoad.Key;
                                break;
                            }
                        }

                    }
                }

                return ownerId;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
