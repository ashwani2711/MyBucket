using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SP.GetCheckOutFileDetails
{
    public class FileStatus
    {
        static StreamWriter writer;
        static TextWriter oldOut = Console.Out;
        string strOutput = null;

        private string objSiteUrl = null;

        private string uName = null;

        private string domain = null;

        private SecureString securePass = null;

        private ClientContext ctx = null;
        private Site site = null;

        private bool IsSPOnline = false;

        private bool UndoCheckOut = false;

        private bool DeleteIfNoPrevVersions = false;

        static bool GetRootWeb = false;
        private string[] WebNames = null;

        /// <summary>
        /// File Status Class
        /// </summary>
        /// <param name="siteUrl">Site Collection URL</param>
        /// <param name="UserName">Username for the SharePoint Online Site Collection</param>
        /// <param name="isonline">Is it SharePoint Online?</param>
        /// <param name="undochkout">Perform Undo Check out? </param>
        /// <param name="deleteifnoversion">Delete the complete file, if no previous versions are found? </param>
        /// <param name="w">Writer Object</param>
        /// <param name="o">Output</param>
        public FileStatus(string siteUrl, string UserName, string dom, bool rWeb, string wnames,
                            bool isonline, bool undochkout, bool deleteifnoversion,
                            StreamWriter w, TextWriter o)
        {
            writer = w;
            oldOut = o;
            objSiteUrl = siteUrl;
            uName = UserName;
            domain = dom;
            GetRootWeb = rWeb;
            WebNames = wnames.Split(',');
            IsSPOnline = isonline;
            UndoCheckOut = undochkout;
            DeleteIfNoPrevVersions = deleteifnoversion;
            strOutput += "File Name\tUrl\tWeb\tWeb Url\tChecked Out to\tEmail" + Environment.NewLine;
        }

        /// <summary>
        /// Processes the Site Collection and iterates all teh subsite
        /// </summary>
        public void ProcessSiteCollection()
        {
            try
            {
                if (UndoCheckOut)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    WriteLine("Performing Undo Check Out");
                    Console.ResetColor();
                }
                WriteLine("Getting the Site Collection Information. " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));

                ctx = new ClientContext(objSiteUrl);
                ctx.RequestTimeout = 36000000;

                if (IsSPOnline)
                {
                    WriteLine("Username: " + uName);
                    securePass = GetPasswordFromText();
                    ctx.Credentials = new SharePointOnlineCredentials(uName, securePass);
                }
                else
                {
                    NetworkCredential networkCred = new NetworkCredential();
                    networkCred.Domain = domain;
                    networkCred.UserName = uName;
                    networkCred.SecurePassword = GetPasswordFromText();

                    ctx.Credentials = networkCred;
                }
                site = ctx.Site;

                Web web = site.RootWeb;

                //Iterates the sub-webs (to generate the reports on the sub-sites)
                //IterateSites(web);
                ProcessWebs(web);

                //Finally write the report to a text file
                if (strOutput != "File Name\tUrl\tWeb\tWeb Url\tChecked Out to\tEmail" + Environment.NewLine)
                {
                    System.IO.File.WriteAllText(Environment.CurrentDirectory + "\\FilesCheckOut-" + DateTime.Now.ToString("dd-MM-yyyy HH mm ss") + ".txt", strOutput);
                }

            }
            catch (Exception ex)
            {
                WriteLine("Error: [ProcessSiteCollection]: " + ex.Message);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        private void ProcessWebs(Web web) 
        {
            if (GetRootWeb)
            {
                ctx.Load(web, rW => rW.Lists, rw => rw.Title, rw => rw.Url, rw => rw.Webs, rw => rw.AppInstanceId);
                ctx.ExecuteQuery();

                //This is checked to ignore SharePoint App webs
                if (web.AppInstanceId.Equals(Guid.Empty))
                {
                    if (UndoCheckOut)
                    {
                        UndoCheckOutForWeb(web);
                    }
                    else
                    {
                        GetFileStatusForWeb(web);
                    }
                }
            }

            if (WebNames.Length > 0)
            {
                ctx.Load(site);
                ctx.ExecuteQuery();

                foreach (string strWebName in WebNames)
                {
                    if (!strWebName.Trim().Equals("'"))
                    {
                        Web subweb = site.OpenWeb(strWebName.Trim());

                        //This is checked to ignore SharePoint App webs
                        IterateSites(subweb);
                    }

                }
            }
        }

        /// <summary>
        /// Iterate Sub Sites (web)
        /// </summary>
        /// <param name="subWeb">Pass the root web object</param>
        public void IterateSites(Web subWeb)
        {
            ctx.Load(subWeb, rW => rW.Lists, rw => rw.Title, rw => rw.Url, rw => rw.Webs, rw => rw.AppInstanceId);
            ctx.ExecuteQuery();

            //This is checked to ignore SharePoint App webs
            if (subWeb.AppInstanceId.Equals(Guid.Empty))
            {
                if (UndoCheckOut)
                {
                    UndoCheckOutForWeb(subWeb);
                }
                else
                {
                    GetFileStatusForWeb(subWeb);
                }
            }

            foreach (Web subweb in subWeb.Webs)
            {
                try
                {
                    ctx.Load(subWeb, rW => rW.Lists, rw => rw.Title, rw => rw.Url, rw => rw.Webs, rw => rw.AppInstanceId);
                    ctx.ExecuteQuery();

                    //Iterate the Sub Sites of the current web
                    IterateSites(subweb);

                    WriteLine(Environment.NewLine);
                }
                catch (Exception ex)
                {
                    WriteLine("Error: [IterateSites]: " + ex.Message);
                }
            }
        }

        public void GetFileStatusForWeb(Web rootWeb)
        {
            WriteLine("Performing the operation on web: " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
            WriteLine("------------------------------------------------");

            //Iterate all the lists within the web
            foreach (List list in rootWeb.Lists)
            {
                try
                {
                    //Load the necessary properties with the list
                    ctx.Load(list, l => l.IsSiteAssetsLibrary, l => l.IsPrivate, l => l.IsCatalog, l => l.IsApplicationList);
                    ctx.ExecuteQuery();

                    //Skip the OOB libraries
                    if (list.BaseType == BaseType.DocumentLibrary && !list.IsApplicationList && !list.IsCatalog
                        && !list.IsPrivate && !list.IsSiteAssetsLibrary &&
                        !list.Title.Equals("VideoLibrary", StringComparison.InvariantCultureIgnoreCase) &&
                            !list.Title.Equals("wfpub", StringComparison.InvariantCultureIgnoreCase) &&
                            !list.Title.Equals("Form Templates", StringComparison.InvariantCultureIgnoreCase))
                    {
                        CamlQuery query = new CamlQuery();
                        //Get only the items where the checkedout user field is not empty (which means the file is checked out)
                        query.ViewXml = "<View Scope=\"RecursiveAll\">" +
                                        "<Query><Where>" +
                                        "<IsNotNull><FieldRef Name='CheckoutUser' />" +
                                        "</IsNotNull></Where></Query></View>";

                        ListItemCollection items = list.GetItems(query);
                        //Load all the items
                        ctx.Load(items);
                        ctx.ExecuteQuery();

                        WriteLine("Found " + items.Count + " item(s) on list " + list.Title);

                        //Iterate each items to get write to the log 
                        foreach (ListItem item in items)
                        {
                            bool userFound = false;
                            string userName = null;
                            string userEmail = string.Empty;

                            try
                            {
                                ctx.Load(item, i => i.File.ServerRelativeUrl, i => i.File.CheckOutType, i => i.File.CheckedOutByUser, i => i.File.Name);
                                ctx.ExecuteQuery();
                                userFound = true;

                                }
                            catch (Exception ex)
                            {
                                if (ex.Message.ToLower().Contains("user cannot be found"))
                                {
                                    ctx.Load(item, i => i.File.ServerRelativeUrl, i => i.File.CheckOutType, i => i.File.Name);
                                    ctx.ExecuteQuery();
                                }
                                else
                                {
                                    WriteLine("Error: " + ex.Message);
                                }
                            }
                            if(userFound)
                            {
                                userName = item.File.CheckedOutByUser.Title;
                                userEmail = item.File.CheckedOutByUser.Email;
                            }
                            else
                            {
                                userName = "Deleted User";
                                userEmail = "Deleted User";
                                // rootWeb.Url + item.File.ServerRelativeUrl + "\t" +
                            }

                            string rootweburl = rootWeb.Url.ToLower().Contains("/sites/") ? rootWeb.Url.Split(new string[] { "/sites/" }, StringSplitOptions.None)[0] : "https://www.avivaworld.com";
                            strOutput += item.File.Name + "\t" +
                              rootweburl + item.File.ServerRelativeUrl + "\t" +
                              rootWeb.Title + "\t" +
                              rootWeb.Url + "\t" +
                              userName + "\t" + userEmail +  Environment.NewLine;

                            WriteLine(item.File.ServerRelativeUrl + ". Check Out Status: " + 
                                item.File.CheckOutType + " Checked Out To: " + userName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteLine("Error writing the output for: " + list.Title);
                    WriteLine("Error: " + ex.Message);
                }
            }
            WriteLine("Completed operation on web " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + Environment.NewLine);
        }

        public void UndoCheckOutForWeb(Web rootWeb)
        {
            WriteLine("Performing the operation on web: " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
            WriteLine("------------------------------------------------");

            //Iterate all the lists within the web
            foreach (List list in rootWeb.Lists)
            {
                try
                {
                    //Load the necessary properties with the list
                    ctx.Load(list, l => l.IsSiteAssetsLibrary, l => l.IsPrivate, l => l.IsCatalog, l => l.IsApplicationList);
                    ctx.ExecuteQuery();

                    //Skip the OOB libraries
                    if (list.BaseType == BaseType.DocumentLibrary && !list.IsApplicationList && !list.IsCatalog
                        && !list.IsPrivate && !list.IsSiteAssetsLibrary &&
                        !list.Title.Equals("VideoLibrary", StringComparison.InvariantCultureIgnoreCase) &&
                            !list.Title.Equals("wfpub", StringComparison.InvariantCultureIgnoreCase) &&
                            !list.Title.Equals("Form Templates", StringComparison.InvariantCultureIgnoreCase))
                    {
                        CamlQuery query = new CamlQuery();
                        //Get only the items where the checkedout user field is not empty (which means the file is checked out)
                        query.ViewXml = "<View Scope=\"RecursiveAll\">" +
                                        "<Query><Where>" +
                                        "<IsNotNull><FieldRef Name='CheckoutUser' />" +
                                        "</IsNotNull></Where></Query></View>";

                        ListItemCollection items = list.GetItems(query);
                        //Load all the items
                        ctx.Load(items);
                        ctx.ExecuteQuery();

                        WriteLine("Found " + items.Count + " item(s) on list " + list.Title);

                        //Iterate each items to get write to the log 
                        foreach (ListItem item in items)
                        {
                            bool userFound = false;
                            try
                            {
                                ctx.Load(item, i => i.File.CheckedOutByUser, i => i.File.ServerRelativeUrl, i => i.File.CheckOutType, i => i.File.Name);
                                ctx.ExecuteQuery();
                                userFound = true;
                            }
                            catch (Exception ex)
                            {
                                //Skip the Checked out user, if the error message is user cannot be found (deleted user)
                                if (ex.Message.ToLower().Contains("user cannot be found"))
                                {
                                    ctx.Load(item, i => i.File.ServerRelativeUrl, i => i.File.CheckOutType, i => i.File.Name);
                                    ctx.ExecuteQuery();
                                    WriteLine("Error: [UndoCheckOutForWeb]: User cannot be found. This error can be ignored");
                                }
                                else
                                {
                                    WriteLine("Error: [UndoCheckOutForWeb]: " + ex.Message);
                                }
                            }
                            try
                            {
                                item.File.UndoCheckOut();
                                ctx.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("You cannot discard check out because there is no checked in version of the document"))
                                {
                                    if (DeleteIfNoPrevVersions)
                                    {
                                        WriteLine("No previous versions found");
                                        WriteLine("Deleting the file. " + rootWeb.Url + item.File.ServerRelativeUrl);
                                        item.DeleteObject();
                                        ctx.ExecuteQuery();
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        WriteLine("Couldn't undo the check out, since no previous versions found. Url: " + rootWeb.Url + item.File.ServerRelativeUrl);
                                        WriteLine("If you want to delete the file completely, set the DeleteIfNoPrevVersions to true in app.config");
                                        Console.ResetColor();
                                    }
                                }
                                else
                                {
                                    WriteLine("Couldn't undo the check out");
                                    WriteLine("Error: [UndoCheckOutForWeb]: " + ex.Message);
                                }
                            }
                            string username = null;
                            string userEmail = string.Empty;
                            if(userFound)
                            {
                                username = item.File.CheckedOutByUser.Title;
                                userEmail = item.File.CheckedOutByUser.Email;
                            }
                            else
                            {
                                username = "Deleted User";
                            }
                            string rootweburl = rootWeb.Url.ToLower().Contains("/sites/") ? rootWeb.Url.Split(new string[] { "/sites/" }, StringSplitOptions.None)[0] : "https://www.avivaworld.com";
                            strOutput += item.File.Name + "\t" +
                                            rootweburl + item.File.ServerRelativeUrl + "\t" +
                                            rootWeb.Title + "\t" +
                                            rootWeb.Url + "\t" +
                                            username + "\t" + userEmail+ Environment.NewLine;
                                            
                            WriteLine(item.File.ServerRelativeUrl + ". Check Out Status: " + item.File.CheckOutType + " Checked Out To: " + username);
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteLine("Error writing the output for: " + list.Title);
                    WriteLine("Error: " + ex.Message);
                }
            }
            WriteLine("Completed operation on web " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + Environment.NewLine);
        }

        #region Utility Methods
        /// <summary>
        /// Writes the Output to Console and also to the file
        /// </summary>
        /// <param name="message">Message to be written to the file and the console</param>
        static void WriteLine(string message)
        {
            Console.SetOut(writer);
            Console.WriteLine(message);

            Console.SetOut(oldOut);
            Console.WriteLine(message);
        }

        #region Get Password From Text
        /// <summary>
        /// Converts the normal string to SecuredString
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public static SecureString GetPasswordFromText()
        {
            Console.Write("Enter your password: ");
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            WriteLine(Environment.NewLine);
            return pwd;
        }
        #endregion

        #endregion


    }



}
