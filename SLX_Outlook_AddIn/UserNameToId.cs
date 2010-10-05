using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sage.SData.Client.Core;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Extensions;


    class UserNameToId
    {
        public static string GetId(string userName)
        {
            try
            {
                ISDataService mydataService = SDataDataService.mydataService();

                SDataResourceCollectionRequest mydataCollection = new SDataResourceCollectionRequest(mydataService);
                mydataCollection.ResourceKind = "Users";
                mydataCollection.QueryValues.Add("where", "UserName eq '" + userName + "'");
                AtomFeed usersFeed = mydataCollection.Read();
                string userId = string.Empty;

                if (usersFeed.Entries.Count() > 0)
                {
                    foreach (AtomEntry entry in usersFeed.Entries)
                    {
                        string tempURI = entry.Id.Uri.AbsoluteUri;
                        userId = tempURI.Substring(tempURI.IndexOf("'") + 1, tempURI.LastIndexOf("'") - tempURI.IndexOf("'") - 1);
                        break;
                    }
                }

                return userId;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }

