using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Data;
using Microsoft.Office.SharePoint;
using Excel = Microsoft.Office.Interop.Excel;
using Bytescout.Spreadsheet;
using System.Diagnostics;
using OfficeOpenXml.Style;
using GemBox.Spreadsheet;

namespace SPAssessment
{

    class SiteData
    {
        ClientContext ClientCntx;
        Web Webpage;
        DataTable Table;
        string Reason;
        bool Status = false;
        string FileSize;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        bool CheckUSer;
        UserCollection SiteUsers;
        public bool GetSiteData(string Url)
        {
            try
            {
                using (ClientCntx = new ClientContext(Url))
                {


                    ClientCntx.Credentials = new SharePointOnlineCredentials(UserCredentials.UserName, UserCredentials.Passwrd);
                    CheckUSer = true;


                    Webpage = ClientCntx.Web;
                    ClientCntx.Load(Webpage);
                    ClientCntx.ExecuteQuery();
                    Console.WriteLine("Share Point Site \n Title: " + Webpage.Title + "; URL: " + Webpage.Url + "; Description: " + Webpage.Description);
                    Console.ReadKey();

                }
            }

            catch (Exception Exceptions)
            {
                CheckUSer = false;
                Console.WriteLine("check user name password");
                WriteToLog.WriteToLogs(Exceptions);
            }
            return CheckUSer;
        }
       
        /****************************Method for Getting document library data and upadting excel sheet************************/
        public void GetDocumentData(string Url)
        {
            try
            {
               
                    string ListName = "MyDocuments";
                    List list = ClientCntx.Web.Lists.GetByTitle(ListName);
                    string fileurl = Url + "/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";

                    File file = ClientCntx.Web.GetFileByUrl(fileurl);
                    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();

                    OpenExcelFile();
                    ClientCntx.Load(file);
                    ClientCntx.ExecuteQuery();
                    using (var pck = new OfficeOpenXml.ExcelPackage())
                    {
                    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                    {
                        if (data != null)
                        {
                            data.Value.CopyTo(mStream);
                            pck.Load(mStream);
                            var ws = pck.Workbook.Worksheets.First();
                            Table = new DataTable();
                            bool hasHeader = true;
                            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                            {
                                Table.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));

                            }
                            var startRow = hasHeader ? 2 : 1;

                            GetUsers();
                            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                            {
                                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                                var row = Table.NewRow();
                                int count = rowNum;
                               
                                foreach (var cell in wsRow)
                                {
                                   
                                        row[cell.Start.Column - 1] = cell.Text;

                                }
                                Status= UpdatLibraryData(row);// upload file int the doccument library 
                                if (Status == true)
                                {
                                    UpdateExcelFile(rowNum, Reason, FileSize, "Success");
                                }
                                else
                                {
                                    UpdateExcelFile(rowNum, Reason, FileSize, "Failed");
                                }
                                Table.Rows.Add(row);
                            }
                            Console.WriteLine('1');
                        }
                    }
                    

                    CloseExcelFile();
                    UploadExcelFile(Url);
                    Console.WriteLine("All Done");
                }

            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while getting the excel data from Sharepoint site");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }

        /**********************************open local excel file and update the data**************************************/
        public void OpenExcelFile()
        {
            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(Settings.LocalFilePath);
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while opening the excel file ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }


        //**********************************save and close local excel file and update the data**************************************//
        public void CloseExcelFile()
        {
            MyBook.Save();
            MyBook.Close();
        }

        /**********************************update local excel file and update the data**************************************/
        private void UpdateExcelFile(int rowNum,string reason,string fileSize,string uploadStatus)
        {
            MySheet.Cells[rowNum,5] = fileSize;
            MySheet.Cells[rowNum, 6] = uploadStatus;
            MySheet.Cells[rowNum, 7] = Reason;
        }


        private bool UpdatLibraryData(DataRow row)
        {
            try
            {
                string Listname = "MyDocuments";
                List list = ClientCntx.Web.Lists.GetByTitle(Listname);
                string FileName = row[Settings.FilePath].ToString();
                string FileStatus = row[Settings.FileStatus].ToString();
                string CreatedBy = row[Settings.CreatedBy].ToString();
                string Department = row[Settings.Department].ToString();

                System.IO.FileInfo Fileinfo = new System.IO.FileInfo(FileName);
                User user = SiteUsers.GetByEmail(CreatedBy); 
                 ClientCntx.Load(user);
                 ClientCntx.ExecuteQuery();

                if (Fileinfo.Exists)
                {

                    double Filesize = (Fileinfo.Length / 1e+6);
                    FileSize = Filesize + "mb";
                    if (Fileinfo.Length < Settings.MaxSIze && Fileinfo.Length > Settings.MinSize)
                    {
                        

                        ListItem DepartmentItem = GetDepartment(Department);
                        FileCreationInformation Fcreateinfo = new FileCreationInformation();
                        Fcreateinfo.Url = Fileinfo.Name;
                        Fcreateinfo.Content = System.IO.File.ReadAllBytes(FileName);
                        Fcreateinfo.Overwrite = true;
                        File FileToUpload = list.RootFolder.Files.Add(Fcreateinfo);
                        ClientCntx.Load(list);
                        ClientCntx.ExecuteQuery();

                        try
                        {
                            ListItem Listitem = FileToUpload.ListItemAllFields;

                            Field field = list.Fields.GetByTitle("FIle_Status");
                            FieldChoice choice = ClientCntx.CastTo<FieldChoice>(field);
                            ClientCntx.Load(choice);
                            ClientCntx.ExecuteQuery();
                            string[] MyStatus = FileStatus.ToUpper().Split(',');
                            string StatusUpload = string.Empty;
                            for (int choicecount = 0; choicecount < MyStatus.Length; choicecount++)
                            {
                                if (choice.Choices.Contains(MyStatus[choicecount].Trim()))
                                {
                                    if (choicecount == MyStatus.Length - 1)
                                    {
                                        StatusUpload = StatusUpload + MyStatus[choicecount];
                                    }
                                    else
                                    {
                                        StatusUpload = StatusUpload + MyStatus[choicecount] + ";";
                                    }
                                }
                            }

                            Listitem["File_Type"] = System.IO.Path.GetExtension(FileName);
                            Listitem["FIle_Status"] = StatusUpload;
                            Listitem["FileCreatedBy"] = user.Title;
                            Listitem["Department_Name"] = DepartmentItem.Id;
                            Listitem.Update();
                            ClientCntx.ExecuteQuery();
                            Reason = "NA";
                            Status = true;
                        }
                        catch (Exception ex)
                        {
                            Reason = "Department not found";
                            Console.WriteLine("Department not found");
                            WriteToLog.WriteToLogs(ex);
                           
                         }

                    }
                    else
                    {
                        if (Fileinfo.Length < Settings.MinSize)
                        {
                            Reason = "File size is Less than Required file size";
                            Console.WriteLine(Reason);
                        }
                        else
                        {
                            Reason = "File size is more than Required file size";
                            Console.WriteLine(Reason);
                        }
                    }
                }
                else
                {
                    Reason ="File Does not exist";
                   
                    Console.WriteLine(Reason);

                }
            }
            catch (Exception ex)
            {
                Reason = "User Does not exist";
                FileSize = "";
                Status = false;
                WriteToLog.WriteToLogs(ex);
            }

            return Status;
        }

        private ListItem GetDepartment(string DepartmentName)
        {
            string Listname = "Department";
            List DeptList = ClientCntx.Web.Lists.GetByTitle(Listname);
            ClientCntx.Load(DeptList);
            ClientCntx.ExecuteQuery();
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Department_Name'/><Value Type = 'Text'>" + DepartmentName.Trim() + "</Value></Eq></Where></Query><RowLimit></RowLimit></View> ";
            ListItemCollection DepartmentItems = DeptList.GetItems(camlQuery);
            ClientCntx.Load(DepartmentItems);
            ClientCntx.ExecuteQuery();
            return DepartmentItems[0];
        }


        /***********************************************Download file from sharepoint site****************************************/
        public void DownloadFile(string Url)
        {
            try
            {
                    string ListName = "Documents";
                    List documentlist = ClientCntx.Web.Lists.GetByTitle(ListName);
                    string urlforworksheet = Url + documentlist.GetItemById(5);
                    var ListItem = documentlist.GetItemById(5);
                    ClientCntx.Load(documentlist);
                    ClientCntx.Load(ListItem, i => i.File);
                    ClientCntx.ExecuteQuery();
                    var FileRef = ListItem.File.ServerRelativeUrl;
                    var FileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ClientCntx, FileRef);
                    //var fileName = System.IO.Path.Combine(urlforworksheet, (string)listItem.File.Name);
                    var FileName = System.IO.Path.Combine(@"D:\", (string)ListItem.File.Name);
                    using (var fileStream = System.IO.File.Create(FileName))
                    {
                        FileInfo.Stream.CopyTo(fileStream);
                    }
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while downloading the file: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }

        /*******************************************upload file again after making changes*****************************************/

        public void UploadExcelFile(string url)
        {
            try
            {
                string Listname = "Documents";

                List list = ClientCntx.Web.Lists.GetByTitle(Listname);
                FileCreationInformation Fcinfo = new FileCreationInformation();
                Fcinfo.Url = "FilePathExcelFile.xlsx";
                Fcinfo.Content = System.IO.File.ReadAllBytes(Settings.LocalFilePath);
                Fcinfo.Overwrite = true;
                File FileToUpload = list.RootFolder.Files.Add(Fcinfo);
                ClientCntx.Load(list);
                ClientCntx.ExecuteQuery();
                Console.WriteLine("Name is : " + Fcinfo.Content);
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while uploading file: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }


        public void GetUsers()
        {
            SiteUsers = ClientCntx.Web.SiteUsers;

            try
            {
                ClientCntx.Load(SiteUsers);
                ClientCntx.ExecuteQuery();
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while getting site users");
                WriteToLog.WriteToLogs(Exceptions);
            }

        }    
        

    }
}

