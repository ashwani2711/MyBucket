using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.WebParts;
using SP.GetWebPartUsage.WebPartPages;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace SP.GetWebPartUsage
{
    public class WebPartUsage
    {
        #region Variables
        static StreamWriter writer;
        static TextWriter oldOut = Console.Out;
       // string strOutput = null;

        private string objSiteUrl = null;

        private string uName = null;
        private string domain = null;

        private SecureString securePass = null;
        private SecureString pwds = null;

        private ClientContext ctx = null;
        private Site site = null;

        private bool IsSPOnline = false;
        SharePointOnlineCredentials cred = null;

        private string[] WebPartsToRetrieve = null;
        static bool GetRootWeb = false;
        private string[] WebNames = null;

        CookieContainer cookieContainer = new CookieContainer();
        #endregion

        public WebPartUsage(string siteUrl, string UserName, string dom,
                            bool isonline,
                            string webparts, bool rWeb, string wnames,
                            StreamWriter w, TextWriter o, SecureString pwd)
        {
            writer = w;
            oldOut = o;
            objSiteUrl = siteUrl;
            uName = UserName;
            domain = dom;
            IsSPOnline = isonline;
            WebPartsToRetrieve = webparts.Split(',');
            GetRootWeb = rWeb;
            pwds = pwd;
            //WebNames = wnames.Split(',');

           
        }

        /// <summary>
        /// Processes the Site Collection and iterates all teh subsite
        /// </summary>
        public void ProcessSiteCollection()
        {
            try
            {
                WriteLine("Getting the Site Collection Information. " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));

                ctx = new ClientContext(objSiteUrl);
                ctx.RequestTimeout = 18000000;

                WriteLine(Environment.NewLine);

                if (IsSPOnline)
                {
                    //Get password for SharePoint Online
                    WriteLine("Getting the detais for SharePoint Online");
                    WriteLine("----------------------------------------");
                    securePass = pwds;
                    cred = new SharePointOnlineCredentials(uName, securePass);
                    ctx.Credentials = cred;
                    SetAuthCookies();
                }
                else
                {
                    //Get password for SharePoint On-Prem/Dedicated as Network Credential
                    WriteLine("Getting the details for SharePoint On-Prem/Dedicated");
                    WriteLine("----------------------------------------------------");
                    NetworkCredential networkCred = new NetworkCredential();
                    networkCred.UserName = uName;
                    networkCred.SecurePassword = pwds;
                    networkCred.Domain = domain;

                    ctx.Credentials = networkCred;
                }
                site = ctx.Site;

                Web web = site.RootWeb;

                //Iterates the sub-webs (to generate the reports on the sub-sites)
                //IterateSites(web);
                ProcessWebs(web);

                //Finally write the report to a text file
                //if (Program.strOutput != "Page Url\tWebPartId\tWeb Part Title\tWeb Part Type\tWebURL\tZoneIndex" + Environment.NewLine)
                //{
                //    System.IO.File.WriteAllText(Environment.CurrentDirectory + "\\WebPartUsage-" + DateTime.Now.ToString("dd-MM-yyyy HH mm ss") + ".txt", Program.strOutput);
                //}

            }
            catch (Exception ex)
            {
                WriteLine("Error: " + ex.Message);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        private void ProcessWebs(Web web)
        {
            try
            {
                if (GetRootWeb)
                {

                    ctx.Load(web, rw => rw.Title, rw => rw.Url, rw => rw.Webs, rw => rw.AppInstanceId, rw => rw.AllProperties, rw => rw.Language);
                    ctx.ExecuteQuery();

                    GetWebPartUsageForWeb(web);
                }
                //swapnil
                if (web.Webs.Count > 0)
                {
                    ctx.Load(site);
                    ctx.ExecuteQuery();

                    foreach (Web w in web.Webs)
                    {
                        //if (!w.webs.Trim().Equals("'"))
                        //{
                        //Web subweb = site.OpenWeb(strWebName.Trim());

                        IterateSites(w);
                        // }

                    }
                }
            }
            catch { Console.WriteLine("Error occured for site {0}",objSiteUrl); }
        }

        /// <summary>
        /// Iterate Sub Sites (web)
        /// </summary>
        /// <param name="rootWeb">Pass the root web object</param>
        public void IterateSites(Web rootWeb)
        {
            ctx.Load(rootWeb, rw => rw.Title, rw => rw.Url, rw => rw.Webs, rw => rw.AppInstanceId, rw => rw.AllProperties, rw=>rw.Language);
            ctx.ExecuteQuery();

            //This is checked to ignore SharePoint App webs
            if (rootWeb.AppInstanceId.Equals(Guid.Empty))
            {
                GetWebPartUsageForWeb(rootWeb);
            }

            foreach (Web subweb in rootWeb.Webs)
            {
                try
                {
                    //Iterate the Sub Sites of the current web
                    IterateSites(subweb);

                    WriteLine(Environment.NewLine);
                }
                catch (Exception ex)
                {
                    WriteLine("Error in web: " + rootWeb.Title);
                    WriteLine(ex.Message);
                }

            }
        }

        public void GetWebPartUsageForWeb(Web rootWeb)
        {
            WriteLine("Performing the operation on web: " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
            WriteLine("------------------------------------------------");
            List list = null;
            string actualWebPartType = null;
            //Iterate all the lists within the web
            try
            {
                string pagesListID = string.Empty;
                string lang = rootWeb.Language.ToString();
                string sitePagesName = string.Empty;
                try
                {
                    pagesListID = rootWeb.AllProperties["__PagesListId"] as string;
                }
                catch(Microsoft.SharePoint.Client.PropertyOrFieldNotInitializedException ex)
                {
                    //do nothing pages library does not exist
                }
                 if(! string.IsNullOrEmpty(pagesListID))
                {
                    list = rootWeb.Lists.GetById(new Guid(pagesListID));
                    GetUsageInList(rootWeb, ref list, ref actualWebPartType);
                }

                switch(lang)
                {
                    case "1036": //fraNCE
                        sitePagesName = "Pages du site";
                        break;
                    case "1040"://italian
                        sitePagesName = "Svetainės puslapiai";
                        break;
                    case "1045"://polish
                        sitePagesName = "Strony witryny";
                        break;
                    case "3082":///spanish
                        sitePagesName = "Páginas del sitio";
                        break;
                    case "1055"://TURKISH
                        sitePagesName = "Site Sayfaları";
                        break;
                    case "1049"://RUSSIAN
                        sitePagesName = "Страницы сайта";
                        break;
                    default: sitePagesName = "Site Pages";
                        break;
                }

                list = rootWeb.Lists.GetByTitle(sitePagesName);
                
                GetUsageInList(rootWeb, ref list, ref actualWebPartType);
            }
            catch (Exception ex)
            {
                WriteLine("Error writing the output for: " + list.Title);
                WriteLine("Error: " + ex.Message);
            }

            WriteLine("Completed operation on web " + rootWeb.Title + ". " + rootWeb.Url + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + Environment.NewLine);
        }

private void GetUsageInList(Web rootWeb, ref List list, ref string actualWebPartType)
{

               

               // list = rootWeb.Lists.GetByTitle("Pages");

                try
                {
                    //Load the necessary properties with the list
                    ctx.Load(list, l => l.Title, l => l.IsSiteAssetsLibrary, l => l.IsPrivate, l => l.IsCatalog, l => l.IsApplicationList);
                    ctx.ExecuteQuery();
                }
                catch
                {
                    list = null;
                }


                //Skip the OOB libraries
                if (list != null)
                {
                    CamlQuery query = new CamlQuery();
                    //Get only the items where the checkedout user field is not empty (which means the file is checked out)
                    query.ViewXml = "<View Scope=\"RecursiveAll\"></View>";

                    ListItemCollection items = list.GetItems(query);
                    //Load all the items
                    ctx.Load(items);
                    ctx.ExecuteQuery();

                    WriteLine("Found " + items.Count + " item(s) on list " + list.Title);

                    //Iterate each items to get write to the log 
                    foreach (ListItem item in items)
                    {
                        ctx.Load(item, i => i.File, i => i.File.ServerRelativeUrl, i => i.File.Name);
                        ctx.ExecuteQuery();

                        if (item.File.Name.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase))
                        {
                            var oFile = item.File;

                            LimitedWebPartManager limitedWebPartManager =
                                oFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

                            ctx.Load(limitedWebPartManager.WebParts,
                              wps => wps.Include(
                              wp => wp.WebPart.Title,
                              wp => wp.Id,
                              wp => wp.WebPart.Properties,
                              wp => wp.WebPart.ZoneIndex));

                            ctx.ExecuteQuery();

                            if (limitedWebPartManager.WebParts.Count == 0)
                            {
                                WriteLine("No Web Parts on this page.");
                            }
                            else
                            {
                                foreach (WebPartDefinition wpDef in limitedWebPartManager.WebParts)
                                {
                                    try
                                    {
                                        string ZoneID = string.Empty;

                                        #region Get Type Name
                                        string webPartPropertiesXml = GetWebPartPropertiesServiceCall(wpDef.Id.ToString(),
                                          oFile.ServerRelativeUrl, rootWeb.Url);

                                        string WebPartTypeName = string.Empty;

                                        if (webPartPropertiesXml.Contains("WebPart/v2"))
                                        {
                                            XmlDataDocument xmldoc = new XmlDataDocument();
                                            xmldoc.LoadXml(webPartPropertiesXml);
                                            WebPartTypeName = xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value;
                                            ZoneID = xmldoc.DocumentElement.GetElementsByTagName("ZoneID").Item(0).FirstChild.Value;
                                            xmldoc = null;
                                        }
                                        else
                                        {
                                            webParts webPartProp = null;
                                            byte[] byteArray = Encoding.UTF8.GetBytes(webPartPropertiesXml);
                                            using (MemoryStream stream = new MemoryStream(byteArray))
                                            {
                                                using (StreamReader streamReader = new StreamReader(stream))
                                                {
                                                    using (System.Xml.XmlReader reader = new XmlTextReader(streamReader))
                                                    {
                                                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                                                        webPartProp = (webParts)serializer.Deserialize(reader);
                                                        WebPartTypeName = webPartProp.webPart.metaData.type.name;
                                                        reader.Close();
                                                        reader.Dispose();
                                                    }
                                                    streamReader.Close();
                                                    streamReader.Dispose();
                                                }
                                                stream.Close();
                                                stream.Flush();
                                                stream.Dispose();
                                            }
                                            byteArray = null;
                                        }

                                        actualWebPartType = GetWebPartShortTypeName(WebPartTypeName);
                                        #endregion

                                        if (WebPartsToRetrieve.Contains(actualWebPartType))
                                        {
                                            WriteLine(oFile.ServerRelativeUrl + "\t" +
                                                        wpDef.Id + "\t" +
                                                        wpDef.Id + "\t" +
                                                        wpDef.WebPart.Title + "\t" +
                                                        actualWebPartType + "\t" +
                                                        rootWeb.Url + "\t" +
                                                        ZoneID + "\t" +
                                                        wpDef.WebPart.ZoneIndex + Environment.NewLine);
                                            Program.strOutput +=
                                                        oFile.ServerRelativeUrl + "," +
                                                        wpDef.Id + "," +
                                                        wpDef.Id + "," +
                                                        wpDef.WebPart.Title + "," +
                                                        actualWebPartType + "," +
                                                        rootWeb.Url + "," +
                                                        ZoneID + "," +
                                                        wpDef.WebPart.ZoneIndex + Environment.NewLine;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteLine("Error in Web Part Definition: " + wpDef.WebPart.Title);
                                        WriteLine(ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
}

        private string GetWebPartID(string webPartID)
        {
            string _webPartID = string.Empty;

            string[] tempStr = webPartID.Split('_');

            if (tempStr.Length > 5)
            {
                _webPartID = webPartID.Remove(0, tempStr[0].Length + 1).Replace('_', '-');
            }
            else
            {
                _webPartID = webPartID.Replace('_', '-');
            }

            return _webPartID;
        }

        private string GetWebPartPropertiesServiceCall(string storageKey, string pageUrl, string weburl)
        {
            string webPartXml = string.Empty;
            try
            {
                var service = new WebPartPages.WebPartPagesWebService();
                service.Url = weburl + "/_vti_bin/webpartpages.asmx";
                service.Timeout = 18000000;

                if (IsSPOnline)
                {
                    service.CookieContainer = cookieContainer;
                }
                else
                {
                    service.Credentials = new  System.Net.NetworkCredential(uName, pwds, domain);
                }

                service.PreAuthenticate = true;

                // Actual web service call which returns the information in string format
                webPartXml = service.GetWebPart2(pageUrl, new Guid(storageKey), Storage.Shared, SPWebServiceBehavior.Version3);
            }
            catch (Exception ex)
            {
                WriteLine("Error in GetWebPartPropertiesServiceCall: " + pageUrl);
                WriteLine(ex.Message);
            }
            return webPartXml;
        }

        private string GetWebPartShortTypeName(string webPartType)
        {
            string _webPartType = string.Empty;

            string[] tempWebPartTypeName = webPartType.Split(',');

            string[] tempWebPartType = tempWebPartTypeName[0].Split('.');
            if (tempWebPartType.Length == 1)
            {
                _webPartType = tempWebPartType[0];
            }
            else
            {
                _webPartType = tempWebPartType[tempWebPartType.Length - 1];
            }
            return _webPartType;
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

       

        public void SetAuthCookies()
        {
            Uri uri = new Uri(objSiteUrl);
            var authCookie = cred.GetAuthenticationCookie(uri);

            cookieContainer.SetCookies(uri, authCookie);
        }
        #endregion
    }
}
