using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sage.SData.Client.Atom;
using Sage.SData.Client.Core;
using Sage.SData.Client.Extensions;

class SDataDataService
{
    public static ISDataService mydataService()
    {
        try
        {
            string sData = SLX_Outlook_AddIn.Properties.Settings.Default.SDATA;
            string userName = SLX_Outlook_AddIn.Properties.Settings.Default.UserName;
            string password = SLX_Outlook_AddIn.Properties.Settings.Default.Password;

            string temp = sData.Substring(sData.Length - 1, 1);

            if (temp == "/")
            {
                sData += "sdata/slx/dynamic/-/";
            }
            else
            {
                sData += "/sdata/slx/dynamic/-/";
            }

            ISDataService service;
            service = new SDataService(sData, userName, password);

            return service;
        }
        catch (Exception ex)
        {
            return null;
        }

    }
}
