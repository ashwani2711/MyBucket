using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.CSV;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.MasterPage
{
    public class MasterPageHelper
    {
       /// <summary>
       /// Initialized of Exception and Logger Class. Deleted the Master Page Replace Usage File
       /// </summary>
       /// <param name="DiscoveryUsage_OutPutFolder"></param>
        public void MasterPage_Initialization(string DiscoveryUsage_OutPutFolder)
        {
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(DiscoveryUsage_OutPutFolder);

            ExceptionCsv.WebApplication = Constants.NotApplicable;
            ExceptionCsv.SiteCollection = Constants.NotApplicable;
            ExceptionCsv.WebUrl = Constants.NotApplicable;

            //Trace Log TXT File Creation Command
            Logger objTraceLogs = Logger.CurrentInstance;
            objTraceLogs.CreateLogFile(DiscoveryUsage_OutPutFolder);
            //Trace Log TXT File Creation Command

            //Delete MasterPage Replace OUTPUT File
            DeleteMasterPage_ReplaceOutPutFiles(DiscoveryUsage_OutPutFolder);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DiscoveryUsage_OutPutFolder"></param>
        /// <param name="New_MasterPageDetails"></param>
        /// <param name="Old_MasterPageDetails"></param>
        /// <param name="SharePointOnline_OR_OnPremise"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="Domain"></param>
        public void ChangeMasterPageForDiscoveryOutPut(string DiscoveryUsage_OutPutFolder, string New_MasterPageDetails = "N/A", string Old_MasterPageDetails = "N/A", string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            try
            {
                //Initialized Exception and Logger. Deleted the Master Page Replace Usage File
                MasterPage_Initialization(DiscoveryUsage_OutPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started : InputCSV ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Started : InputCSV ##############");
               
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + DiscoveryUsage_OutPutFolder);
                Console.WriteLine("[ChangeMasterPageForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + DiscoveryUsage_OutPutFolder);
                
                //Reading Master Page Input File
                IEnumerable<MasterPageInput> objMpInput;
                ReadMasterPagesFromDiscoveryUsageFiles(DiscoveryUsage_OutPutFolder, out objMpInput);
                
                //Changing Master Pages for INPUT CSV
                ChangeMasterPage_UsingDiscoveryOutPut(objMpInput, DiscoveryUsage_OutPutFolder, New_MasterPageDetails, Old_MasterPageDetails,SharePointOnline_OR_OnPremise,UserName,Password,Domain);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                Console.WriteLine("[END] EXIT FROM FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed : InputCSV ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed : InputCSV ##############");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[Exception] FUNCTION ChangeMasterPageForDiscoveryOutPut. Exception Message: " + ex.Message);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ChangeMasterPageForDiscoveryOutPut. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForDiscoveryOutPut", ex.GetType().ToString());
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="outPutFolder"></param>
        /// <param name="objMpInput"></param>
        /// <param name="New_MasterPageDetails"></param>
        /// <param name="Old_MasterPageDetails"></param>
        public void ReadMasterPagesFromDiscoveryUsageFiles(string outPutFolder, out IEnumerable<MasterPageInput> objMpInput, string New_MasterPageDetails = "N/A", string Old_MasterPageDetails = "N/A")
        {
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ReadMasterPagesFromDiscoveryUsageFiles");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadMasterPagesFromDiscoveryUsageFiles] [START] Calling function ImportCsv.Read<MasterPageInput>. Master Page Input CSV file is available at " + outPutFolder + " and Master Page Input file name is " + Constants.MasterPageInput);
            
            objMpInput = null;
            objMpInput = ImportCsv.Read<MasterPageInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.MasterPageInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadMasterPagesFromDiscoveryUsageFiles] [END] Read all the INPUT from Master Page and saved in List - out IEnumerable<MasterPageInput> objMpInput, for processing.");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ::: ReadMasterPagesFromDiscoveryUsageFiles");
        }

        /// <summary>
        /// This Function will take Discovery - Master Page Usage Reports as a Input. 
        /// It will read all the rows and Update the Master Page for all records/sites
        /// </summary>
        /// <param name="objMPInput"></param>
        /// <param name="outPutFolder"></param>
        /// <param name="New_MasterPageDetails"></param>
        /// <param name="Old_MasterPageDetails"></param>
        /// <param name="SharePointOnline_OR_OnPremise"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="Domain"></param>
        public void ChangeMasterPage_UsingDiscoveryOutPut(IEnumerable<MasterPageInput> objMPInput, string outPutFolder, string New_MasterPageDetails = "N/A", string Old_MasterPageDetails = "N/A", string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            
            List<MasterPageBase> _WriteMasterList = new List<MasterPageBase>();

            try
            {
                if (objMPInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ChangeMasterPage_UsingDiscoveryOutPut");
                    bool headerMasterPage = false;
                    foreach (MasterPageInput objInput in objMPInput)
                    {
                        //This is for Exception Comments:
                        exceptionCommentsInfo1 = "<Input>New MasterPage Url = " + New_MasterPageDetails + ", <Input> OLD MasterUrl: " + Old_MasterPageDetails + ", WebUrl: " + objInput.WebUrl + ", CustomMasterUrlStatus" + objInput.CustomMasterUrlStatus + "MasterUrlStatus" + objInput.MasterUrlStatus;
                        //This is for Exception Comments:
                        
                        MasterPageBase objMPBase = null;
                        //!= "all" = This will update the master page only in those sites/webs, which have Custom and Master URL == <Input>Old_MasterPageDetails
                        if (Old_MasterPageDetails.ToLower().Trim() != "all" && Old_MasterPageDetails.ToLower().Trim() != "" )
                        {
                            objMPBase = new MasterPageBase();
                            objMPBase = ChangeMasterPageForDiscoveryUsage(objInput.WebUrl, New_MasterPageDetails, Old_MasterPageDetails, Old_MasterPageDetails, Convert.ToBoolean(objInput.CustomMasterUrlStatus), Convert.ToBoolean(objInput.MasterUrlStatus), SharePointOnline_OR_OnPremise, UserName, Password, Domain); 
                        }
                        //all = This will update the master page from all input web/site Using New_MasterPageDetails
                        else 
                        {
                            objMPBase = new MasterPageBase();
                            objMPBase = ChangeMasterPageForDiscoveryUsage(objInput.WebUrl, New_MasterPageDetails, objInput.MasterUrl, objInput.CustomMasterUrl, Convert.ToBoolean(objInput.CustomMasterUrlStatus), Convert.ToBoolean(objInput.MasterUrlStatus), SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                        }

                        if (objMPBase != null)
                        { _WriteMasterList.Add(objMPBase); }
                    }

                    FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.MasterPageUsage, ref _WriteMasterList,
                    ref headerMasterPage);

                    Console.WriteLine("[END] EXIT FROM FUNCTION ChangeMasterPage_UsingDiscoveryOutPut");
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ChangeMasterPage_UsingDiscoveryOutPut");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[Exception] FUNCTION ChangeMasterPage_UsingDiscoveryOutPut. Exception Message:" + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPage_UsingDiscoveryOutPut", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ChangeMasterPage_UsingDiscoveryOutPut. Exception Message:" + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="WebUrl"></param>
        /// <param name="NewMasterPageURL"></param>
        /// <param name="OldMasterPageURL"></param>
        /// <param name="OldCustomMasterPageURL"></param>
        /// <param name="CustomMasterUrlStatus"></param>
        /// <param name="MasterUrlStatus"></param>
        /// <param name="SharePointOnline_OR_OnPremise"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="Domain"></param>
        /// <returns></returns>
        public MasterPageBase ChangeMasterPageForDiscoveryUsage(string WebUrl, string NewMasterPageURL, string OldMasterPageURL = "N/A", string OldCustomMasterPageURL = "N/A", bool CustomMasterUrlStatus = true, bool MasterUrlStatus = true, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            MasterPageBase objMaster = null;
            ExceptionCsv.WebUrl = WebUrl;
            
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ChangeMasterPageForDiscoveryUsage for WebUrl: " + WebUrl);
            Console.WriteLine("[START] ENTERING IN FUNCTION ChangeMasterPageForDiscoveryUsage for WebUrl: " + WebUrl);

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == "OP")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForDiscoveryUsage] ENTERING IN FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForDiscoveryUsage] EXIT FROM FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == "OL")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForDiscoveryUsage] ENTERING IN FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForDiscoveryUsage] EXIT FROM FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext!=null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    //Create NEW Master Page Relative URL
                    string _strNewMasterPageUrl = string.Empty;
                    _strNewMasterPageUrl = GetMasterPageRelativeURL(clientContext, NewMasterPageURL);

                    //Create OldMasterPageURL Relative URL
                    string _strOldMasterPageURL = string.Empty;
                    _strOldMasterPageURL = GetMasterPageRelativeURL(clientContext, OldMasterPageURL);

                    //Create OldCustomMasterPageURL Relative URL
                    string _strOldCustomMasterPageURL = string.Empty;
                    _strOldCustomMasterPageURL = GetMasterPageRelativeURL(clientContext, OldCustomMasterPageURL);

                    exceptionCommentsInfo1 = "OldMasterPageURL:" + _strOldMasterPageURL + ", OldCustomMasterPageURL:" + _strOldCustomMasterPageURL + " New Master URL: " + _strNewMasterPageUrl + " CustomMasterUrlStatus: " + CustomMasterUrlStatus + ", MasterUrlStatus: " + MasterUrlStatus;

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + _strNewMasterPageUrl);
                    Console.WriteLine("[ChangeMasterPageForDiscoveryUsage]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + _strNewMasterPageUrl);

                    //Check if new master page is available in Gallery
                    if (Check_MasterPageExistsINGallery(clientContext, _strNewMasterPageUrl))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[Check_MasterPageExistsINGallery]: This New Master Page is present in Gallery: " + _strNewMasterPageUrl);
                        Console.WriteLine("[Check_MasterPageExistsINGallery]: This New Master Page is present in Gallery: " + _strNewMasterPageUrl);

                        bool _UpdateMasterPage = false;

                        if (CustomMasterUrlStatus && _strOldCustomMasterPageURL.ToLower().Trim() == web.CustomMasterUrl.ToString().Trim().ToLower())
                        { 
                            web.CustomMasterUrl = _strNewMasterPageUrl; 
                            _UpdateMasterPage = true;

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage]: Updated Custom Master Page " + _strOldCustomMasterPageURL + " with new Master Page URL " + _strNewMasterPageUrl);
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage]: Updated Custom Master Page " + _strOldCustomMasterPageURL + " with new Master Page URL " + _strNewMasterPageUrl);
                        }
                        else 
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldCustomMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldCustomMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                        }

                        if (MasterUrlStatus && _strOldMasterPageURL.ToLower().Trim() == web.MasterUrl.ToString().Trim().ToLower())
                        { 
                            web.MasterUrl = _strNewMasterPageUrl;
                            _UpdateMasterPage = true;

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + _strNewMasterPageUrl);
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + _strNewMasterPageUrl);
                        } 
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage]:  [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage]:  [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                        }
                        
                        if (_UpdateMasterPage)
                        {
                            objMaster = new MasterPageBase();

                            web.Update();

                            clientContext.Load(web);
                            clientContext.ExecuteQuery();

                            //Added
                            objMaster.CustomMasterUrl = web.CustomMasterUrl;
                            objMaster.MasterUrl = web.MasterUrl;
                            objMaster.WebApplication = Constants.NotApplicable;
                            objMaster.SiteCollection = Constants.NotApplicable;
                            objMaster.WebUrl = web.Url;
                            objMaster.OLD_CustomMasterUrl = _strOldCustomMasterPageURL;
                            objMaster.OLD_MasterUrl = _strOldMasterPageURL;
                            //Added

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage] Changed Master Page for - " + WebUrl + ", New Master Page is " + _strNewMasterPageUrl);
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage] Changed Master Page for - " + WebUrl + ", New Master Page is " + _strNewMasterPageUrl);
                        }
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage] The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                            Console.WriteLine("[ChangeMasterPageForDiscoveryUsage] The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                        }
                    }
                    else
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage] We have not changed the master page because this new Master Page " + _strNewMasterPageUrl + " is not present in Gallary, for Web" + WebUrl);
                        Console.WriteLine("[ChangeMasterPageForDiscoveryUsage] We have not changed the master page because this new Master Page " + _strNewMasterPageUrl + " is not present in Gallary, for Web" + WebUrl);
                    }
                }
                else
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryUsage] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                    Console.WriteLine("[ChangeMasterPageForDiscoveryUsage] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION [ChangeMasterPageForDiscoveryUsage] for WebUrl: " + WebUrl);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "**");
                Console.WriteLine("[END] EXIT FROM FUNCTION [ChangeMasterPageForDiscoveryUsage] for WebUrl: " + WebUrl);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EXCEPTION] [ChangeMasterPageForDiscoveryUsage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangeMasterPageForDiscoveryUsage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForDiscoveryUsage", ex.GetType().ToString(), exceptionCommentsInfo1);   
            }
            return objMaster;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="outPutFolder"></param>
        /// <param name="SiteCollectionUrl"></param>
        /// <param name="NewMasterPageURL"></param>
        /// <param name="OldMasterPageURL"></param>
        /// <param name="CustomMasterUrlStatus"></param>
        /// <param name="MasterUrlStatus"></param>
        /// <param name="SharePointOnline_OR_OnPremise"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="Domain"></param>
        public void ChangeMasterPageForSiteCollection(string outPutFolder, string SiteCollectionUrl, string NewMasterPageURL, string OldMasterPageURL = "N/A", bool CustomMasterUrlStatus = true, bool MasterUrlStatus = true, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            List<MasterPageBase> _WriteMasterList = new List<MasterPageBase>();
            //Initialized Exception and Logger. Deleted the Master Page Replace Usage File

            MasterPage_Initialization(outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started - For Site Collection ##############");
            Console.WriteLine("############## Master Page Trasnformation Utility Execution Started - For Site Collection ##############");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangeMasterPageForSiteCollection");
            Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangeMasterPageForSiteCollection");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);
            Console.WriteLine("[ChangeMasterPageForSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);
            Console.WriteLine("[ChangeMasterPageForSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);
                
            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                
                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == "OP")
                {
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(SiteCollectionUrl, UserName, Password, Domain);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == "OL")
                {
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(SiteCollectionUrl, UserName, Password);
                }

                if (clientContext != null)
                {
                    bool headerMasterPage = false;
                    MasterPageBase objMPBase = new MasterPageBase();
                    
                    Web rootWeb = clientContext.Web;
                    clientContext.Load(rootWeb);
                    clientContext.ExecuteQuery();

                    //Root Web
                    objMPBase = ChangeMasterPageForWeb(outPutFolder, rootWeb.Url.ToString(), NewMasterPageURL, OldMasterPageURL, CustomMasterUrlStatus, MasterUrlStatus, "SiteCollection", SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                    if (objMPBase != null)
                    {
                        _WriteMasterList.Add(objMPBase);
                    }
                    WebCollection webCollection = rootWeb.Webs;      
                    clientContext.Load(webCollection);
                    clientContext.ExecuteQuery();

                    ExceptionCsv.SiteCollection = SiteCollectionUrl;

                    foreach (Web webSite in webCollection)
                    {
                        try
                        {
                            //Web
                            objMPBase = ChangeMasterPageForWeb(outPutFolder, webSite.Url, NewMasterPageURL, OldMasterPageURL, CustomMasterUrlStatus, MasterUrlStatus, "SiteCollection", SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                            
                            if (objMPBase != null)
                            { _WriteMasterList.Add(objMPBase); }
                        }
                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangeMasterPageForSiteCollection] ChangeMasterPageForSiteCollection. Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                            Console.WriteLine("[EXCEPTION] [ChangeMasterPageForSiteCollection] Exception Message: " + ex.Message);
                        }
                    }

                    FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.MasterPageUsage, ref _WriteMasterList,
                    ref headerMasterPage);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] [ChangeMasterPageForSiteCollection] EXIT FROM FUNCTION ChangeMasterPageForSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);
                Console.WriteLine("[END] [ChangeMasterPageForSiteCollection] EXIT FROM FUNCTION ChangeMasterPageForSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed - For Site Collection ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed - For Site Collection ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangeMasterPageForSiteCollection] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.WriteLine("[EXCEPTION] [ChangeMasterPageForSiteCollection] Exception Message: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="outPutFolder"></param>
        /// <param name="WebUrl"></param>
        /// <param name="NewMasterPageURL"></param>
        /// <param name="OldMasterPageURL"></param>
        /// <param name="CustomMasterUrlStatus"></param>
        /// <param name="MasterUrlStatus"></param>
        /// <param name="ActionType"></param>
        /// <param name="SharePointOnline_OR_OnPremise"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        /// <param name="Domain"></param>
        /// <returns></returns>
        public MasterPageBase ChangeMasterPageForWeb(string outPutFolder, string WebUrl, string NewMasterPageURL, string OldMasterPageURL = "N/A", bool CustomMasterUrlStatus = true, bool MasterUrlStatus = true, string ActionType = "", string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            bool headerMasterPage = false;
            List<MasterPageBase> _WriteMasterList = null;
            ExceptionCsv.WebUrl = WebUrl;

            ///<ActionType==""> That means this function running only for a web. We have to write the output in this function only
            ///<Action Type=="SiteCollection"> The function will return object MasterPageBase, and consolidated output will be written in SiteCollection function - ChangeMasterPageForSiteCollection
            
            if(ActionType=="")
            {
                MasterPage_Initialization(outPutFolder);
                _WriteMasterList = new List<MasterPageBase>();

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangeMasterPageForWeb");
                Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangeMasterPageForWeb");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ChangeMasterPageForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] WebUrl is " + WebUrl);
                Console.WriteLine("[ChangeMasterPageForWeb] WebUrl is " + WebUrl);
            }
            
            string exceptionCommentsInfo1 = string.Empty;
            MasterPageBase objMaster = new MasterPageBase();
            
            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == "OP")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == "OL")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                    Console.WriteLine("[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                    Web web = clientContext.Web;
                    
                    //Load Web to get old Master Page details
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();
                    //Load Web to get old Master Page details

                    //Create New Master Page Relative URL
                    string masterPageUrl = string.Empty;
                    masterPageUrl = GetMasterPageRelativeURL(clientContext, NewMasterPageURL);

                    //Create OldMasterPageURL Relative URL
                    string _strOldMasterPageURL = string.Empty;
                    if (OldMasterPageURL.Trim().ToLower() != "" && OldMasterPageURL.Trim().ToLower() != "all")
                    {
                        _strOldMasterPageURL = GetMasterPageRelativeURL(clientContext, OldMasterPageURL);
                    }

                    //Prepare Exception Comments
                    exceptionCommentsInfo1 = "New Master URL: " + masterPageUrl + ", OldMasterPageURL="+_strOldMasterPageURL+", CustomMasterUrlStatus: " + CustomMasterUrlStatus + ", MasterUrlStatus: " + MasterUrlStatus;

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + masterPageUrl);
                    Console.WriteLine("[ChangeMasterPageForWeb]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + masterPageUrl);
                    
                    //Check if new master page is available in Gallery
                    if (Check_MasterPageExistsINGallery(clientContext, masterPageUrl))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Check_MasterPageExistsINGallery: This New Master Page is present in Gallery: " + masterPageUrl);
                        Console.WriteLine("[ChangeMasterPageForWeb] Check_MasterPageExistsINGallery: This New Master Page is present in Gallery: " + masterPageUrl);

                        //Added in Output Object <objMaster> - To Write old Master Page details
                        objMaster.OLD_CustomMasterUrl = web.CustomMasterUrl;
                        objMaster.OLD_MasterUrl = web.MasterUrl;
                        //Added in Output Object <objMaster> - To Write old Master Page details

                        if (OldMasterPageURL.Trim().ToLower() != "" && OldMasterPageURL.Trim().ToLower() != "all")
                        {
                            bool _UpdateMasterPage = false;

                            if (CustomMasterUrlStatus && _strOldMasterPageURL.ToLower().Trim() == web.CustomMasterUrl.ToString().Trim().ToLower())
                            {
                                web.CustomMasterUrl = masterPageUrl;
                                _UpdateMasterPage = true;

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Updated Custom Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: Updated Custom Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                            }

                            if (MasterUrlStatus && _strOldMasterPageURL.ToLower().Trim() == web.MasterUrl.ToString().Trim().ToLower())
                            {
                                web.MasterUrl = masterPageUrl;
                                _UpdateMasterPage = true;

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                            }

                            if (_UpdateMasterPage)
                            {
                                web.Update();

                                clientContext.Load(web);
                                clientContext.ExecuteQuery();

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                            }
                        }
                        else
                        {
                            if (CustomMasterUrlStatus)
                            { web.CustomMasterUrl = masterPageUrl; }

                            if (MasterUrlStatus)
                            { web.MasterUrl = masterPageUrl; }

                            //Update Web
                            web.Update();

                            //Load Web to get Updated Details
                            clientContext.Load(web);
                            clientContext.ExecuteQuery();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL ==\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                            Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL ==\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                        }

                        //Added in Output Object <objMaster> 
                        objMaster.CustomMasterUrl = web.CustomMasterUrl;
                        objMaster.MasterUrl = web.MasterUrl;
                        objMaster.WebApplication = Constants.NotApplicable;
                        objMaster.SiteCollection = Constants.NotApplicable;
                        objMaster.WebUrl = web.Url;
                        //Added in Output Object <objMaster> 
                    }
                    else 
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] We have not changed the master page because this new Master Page " + masterPageUrl + " is not present in Gallary, for Web " + WebUrl);
                        Console.WriteLine("[ChangeMasterPageForWeb] We have not changed the master page because this new Master Page " + masterPageUrl + " is not present in Gallary, for Web " + WebUrl);
                    }
                }
                else
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                    Console.WriteLine("[ChangeMasterPageForWeb] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] [ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "**");
                Console.WriteLine("[END] [ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                Console.WriteLine("**");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][ChangeMasterPageForWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.WriteLine("[EXCEPTION][ChangeMasterPageForWeb] Exception Message: " + ex.Message + " for Web:  " + WebUrl);
            }

            if (ActionType == "")
            {
                if (objMaster != null)
                {
                    _WriteMasterList.Add(objMaster);
                }

                FileUtility.WriteCsVintoFile(outPutFolder +@"\" + Constants.MasterPageUsage, ref _WriteMasterList,
                        ref headerMasterPage);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Writing the Replace Output CSV file after replacing the master page - FileUtility.WriteCsVintoFile");
                Console.WriteLine("[ChangeMasterPageForWeb] Writing the Replace Output CSV file after replacing the master page - FileUtility.WriteCsVintoFile");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                Console.WriteLine("[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed for Web ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed  for Web ##############");
            }

            return objMaster;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="MasterPageURL"></param>
        /// <returns></returns>
        public string GetMasterPageRelativeURL(ClientContext clientContext, string MasterPageURL)
        {
            string _masterPageUrl = string.Empty;

            //User has Input Master Page URL from Root Gallery
            if (MasterPageURL.ToLower().StartsWith("/_catalogs/masterpage/"))
            {
                Web rootWeb = clientContext.Site.RootWeb;
                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + MasterPageURL;
            }
            else if (MasterPageURL.ToLower().Contains("/_catalogs/masterpage/"))
            {
                _masterPageUrl = MasterPageURL;
            }
            else
            {
                Web rootWeb = clientContext.Site.RootWeb;

                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                if (rootWeb.ServerRelativeUrl.ToString().EndsWith("/"))
                {
                    _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + "_catalogs/masterpage/" + MasterPageURL;
                }
                else
                {
                    _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + "/_catalogs/masterpage/" + MasterPageURL;
                }
            }

            return _masterPageUrl;

        }

        /// <summary>
        /// How to determine if a file exists in a SharePoint SPFolder
        /// CSOM: File Check in SP Gallary. It would actually throw an exception if the file doesn't exist
        /// </summary>
        /// <param name="WebUrl"></param>
        /// <param name="MasterPageURL"></param>
        /// <returns></returns>
        public bool Check_MasterPageExistsINGallery(string WebUrl, string MasterPageURL)
        {
            using (var clientContext = new ClientContext(WebUrl))
            {
                Web web = clientContext.Web;
                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(MasterPageURL);
                bool bExists = false;

                try
                {
                    clientContext.Load(file);
                    clientContext.ExecuteQuery(); //Raises exception if the file doesn't exist
                    bExists = file.Exists;  //may not be needed - here for good measure
                }
                catch { }

                return bExists;
            }
        }

       /// <summary>
       /// Master Pages and Page Layouts Always be Saved in Root Site inside the "_catalogs" folder
       /// Function Used For: How to determine if a file exists in a SharePoint SPFolder - True => Exists, False => Not Exists
       /// CSOM: File Check in SP Gallary. It would actually throw an exception if the file doesn't exist 
       /// </summary>
       /// <param name="clientContext"></param>
       /// <param name="MasterPageURL"></param>
       /// <returns></returns>
        public bool Check_MasterPageExistsINGallery(ClientContext clientContext, string MasterPageURL)
        {
            //Checking The File in Root Web Gallery
            Microsoft.SharePoint.Client.File file = clientContext.Site.RootWeb.GetFileByServerRelativeUrl(MasterPageURL);

            bool bExists = false;
            try
            {
                clientContext.Load(file);
                clientContext.ExecuteQuery(); //Raises exception if the file doesn't exist
                bExists = file.Exists; 
            }
            catch { }

            return bExists;
        }
        
        /// <summary>
        /// This function delete all the existing files from <outPutFolder> folder
        /// </summary>
        /// <param name="outPutFolder"></param>
        private void DeleteMasterPage_ReplaceOutPutFiles(string outPutFolder)
        {
            FileUtility.DeleteFiles(outPutFolder + @"\" + Constants.MasterPageUsage);
        }
        private string GetPageNameFromURL(string URL)
        {
            string FileName = string.Empty;

            if (URL != null)
            { FileName = System.IO.Path.GetFileName(URL); }

            return FileName;
        }
        private string GetPageNameWithSuffix(string PageNameWithExtension, string Suffix)
        {
            string PageNameWithSuffix = string.Empty;

            string Name = System.IO.Path.GetFileNameWithoutExtension(PageNameWithSuffix);
            string Extension = System.IO.Path.GetExtension(PageNameWithSuffix);

            return PageNameWithSuffix = Name + Suffix + Extension;
        }
        
    }
}
