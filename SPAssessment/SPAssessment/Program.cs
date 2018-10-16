using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
namespace SPAssessment
{
    class Program
    {
        static bool LoginStatus;
        static void Main(string[] args)
        {
            UserCredentials.Getdata();
            string Url = "https://acuvatehyd.sharepoint.com/teams/shubhamtrial";
            SiteData sitedata = new SiteData();
            LoginStatus=sitedata.GetSiteData(Url);
            if (LoginStatus == true)
            {
                sitedata.DownloadFile(Url);
                sitedata.GetDocumentData(Url);
                //sitedata.UploadData(Url, UserName, passwrd);
            }
            else
            {
                Main(args);
            }
            Console.ReadKey();
        }
    }
}
