using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.Utilities;
using System.Xml;
using Transformation.PowerShell.Common.CSV;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    public class SiteColumnAndContentTypeHelper
    {

        #region Common Functions
        private void SiteColumnAndContentType_Initialization(string DiscoveryUsage_OutPutFolder, string Type)
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

            if (Type == "DUPLICATE-SITE-COLUMN")
            {
                //SiteColumns Replace OUTPUT File
                FileUtility.DeleteFiles(DiscoveryUsage_OutPutFolder + @"\" + Constants.SiteColumnDuplicateOutput);
            }
            else if (Type == "ADD-SITE-COLUMN-IN-CONTENT-TYPE")
            {
                //SiteColumns Replace OUTPUT File
                FileUtility.DeleteFiles(DiscoveryUsage_OutPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput);
            }
            else if (Type == "REPLACE-CONTENT-TYPE-IN-LIST")
            {
                //SiteColumns Replace OUTPUT File
                FileUtility.DeleteFiles(DiscoveryUsage_OutPutFolder + @"\" + Constants.ContentType_Add_To_ListOutput);
            }
            else if (Type == "DUPLICATE-CONTENT-TYPE")
            {
                //SiteColumns Replace OUTPUT File
                FileUtility.DeleteFiles(DiscoveryUsage_OutPutFolder + @"\" + Constants.ContentTypeDuplicateOutput);
            }
        }
        public ContentType GetContentTypeByName(ClientContext clientContext, Web web, string ContentTypeName, bool IsRoot = false)
        {
            ContentTypeCollection contentTypes = null;

            if (IsRoot)
            { contentTypes = clientContext.Site.RootWeb.ContentTypes; }
            else
            { contentTypes = web.ContentTypes; }

            clientContext.Load(contentTypes);
            clientContext.ExecuteQuery();

            return contentTypes.FirstOrDefault(o => o.Name == ContentTypeName);
        }
        public ContentType GetContentTypeByName(ClientContext clientContext, string ContentTypeName, bool IsRoot = false)
        {
            ContentTypeCollection contentTypes = null;

            if (IsRoot)
            { contentTypes = clientContext.Site.RootWeb.ContentTypes; }
            else
            { contentTypes = clientContext.Web.ContentTypes; }

            clientContext.Load(contentTypes);
            clientContext.ExecuteQuery();

            return contentTypes.FirstOrDefault(o => o.Name == ContentTypeName);
        }
        public static ContentType GetContentTypeByID(ClientContext clientContext, Web web, string ContentTypeID, bool IsRoot = false)
        {
            ContentType contentType = null;

            if (IsRoot)
            {
                contentType = clientContext.Site.RootWeb.ContentTypes.GetById(ContentTypeID);
            }
            else
            { contentType = clientContext.Web.ContentTypes.GetById(ContentTypeID); }

            return contentType;
        }
        public bool Check_ValidateAllInputFields(string outPutFolder, string WebUrl, string oldSiteColumn_InternalName, string oldSiteColumn_ID, string newSiteColumn_InternalName, string newSiteColumn_DisplayName, string UserName, string Password)
        {
            bool _Check = true;

            if (outPutFolder == "" || outPutFolder == null)
            {
                _Check = false;
                Console.WriteLine("outPutFolder details cannot be blank!");
                return _Check;
            }
            else if (WebUrl == "" || WebUrl == null)
            {
                _Check = false;
                Console.WriteLine("WebUrl field cannot be blank!");
                return _Check;
            }
            else if (oldSiteColumn_InternalName == "" || oldSiteColumn_InternalName == null)
            {
                _Check = false;
                Console.WriteLine("oldSiteColumn_InternalName field cannot be blank!");
                return _Check;
            }
            else if (newSiteColumn_InternalName == "" || newSiteColumn_InternalName == null)
            {
                _Check = false;
                Console.WriteLine("newSiteColumn_InternalName field cannot be blank!");
                return _Check;
            }
            else if (newSiteColumn_DisplayName == "" || newSiteColumn_DisplayName == null)
            {
                _Check = false;
                Console.WriteLine("newSiteColumn_DisplayName field cannot be blank!");
                return _Check;
            }
            else if (UserName == "" || UserName == null)
            {
                _Check = false;
                Console.WriteLine("UserName field cannot be blank!");
                return _Check;
            }
            else if (Password == "" || Password == null)
            {
                _Check = false;
                Console.WriteLine("Password field cannot be blank!");
                return _Check;
            }

            return _Check;
        }
        public dynamic isExist_Helper(ClientContext context, String fieldToCheck, String type)
        {
            var isExist = 0;
            ListCollection listCollection = context.Web.Lists;
            ContentTypeCollection cntCollection = context.Web.ContentTypes;
            FieldCollection fldCollection = context.Web.Fields;
            switch (type)
            {
                case "list":
                    context.Load(listCollection, lsts => lsts.Include(list => list.Title).Where(list => list.Title == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = listCollection.Count;
                    break;
                case "contenttype":
                    context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name).Where(ct => ct.Name == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = cntCollection.Count;
                    break;
                case "contenttypeName":
                    context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name, ct => ct.Id).Where(ct => ct.Name == fieldToCheck));
                    context.ExecuteQuery();
                    foreach (ContentType ct in cntCollection)
                    {
                        return ct.Id.ToString();
                    }
                    break;
                case "field":
                    context.Load(fldCollection, fld => fld.Include(ft => ft.Title).Where(ft => ft.Title == fieldToCheck));
                    try
                    {
                        context.ExecuteQuery();
                        isExist = fldCollection.Count;
                    }
                    catch (Exception e)
                    {
                        if (e.Message == "Unknown Error")
                        {
                            isExist = fldCollection.Count;
                        }
                    }
                    break;
                case "listcntype":
                    List lst = context.Web.Lists.GetByTitle(fieldToCheck);
                    ContentTypeCollection lstcntype = lst.ContentTypes;
                    context.Load(lstcntype, lstc => lstc.Include(lc => lc.Name).Where(lc => lc.Name == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = lstcntype.Count;
                    break;
            }
            return isExist;
        }
        #endregion

        #region SiteColumns Creation
        public void SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV(string oldSiteColumn_InternalName, string oldSiteColumn_ID, string newSiteColumn_InternalName, string newSiteColumn_DisplayName, string SiteColumnUsageFilePath, string OutPutFolder, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                SiteColumnAndContentType_Initialization(OutPutFolder, "DUPLICATE-SITE-COLUMN");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Site-Columns/Custom-Fields Trasnformation Utility Execution Started - Using Site Columns Or Custom Fields Input CSV ##############");
                Console.WriteLine("############## Site-Columns/Custom-Fields Trasnformation Utility Execution Started - Using Site Columns Or Custom Fields Input CSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV");
                Console.WriteLine("[START] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV");
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + OutPutFolder);
                Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + OutPutFolder);

                //Reading Input File
                IEnumerable<SiteColumnInput> objSCInput;
                SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV(oldSiteColumn_InternalName, oldSiteColumn_ID, newSiteColumn_InternalName, newSiteColumn_DisplayName, SiteColumnUsageFilePath, OutPutFolder, out objSCInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                //Reading Input File

                Console.WriteLine("[END] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Site-Columns/Custom-Fields Trasnformation Utility Execution Completed - Using Site Columns Or Custom Fields Input CSV ##############");
                Console.WriteLine("############## Site-Columns/Custom-Fields Trasnformation Utility Execution Completed - Using Site Columns Or Custom Fields Input CSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "SiteColumnCSOMCreation", ex.Message, ex.ToString(), "SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private void SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV(string oldSiteColumn_InternalName, string oldSiteColumn_ID, string newSiteColumn_InternalName, string newSiteColumn_DisplayName, string SiteColumnUsageFilePath, string outPutFolder, out IEnumerable<SiteColumnInput> objSCInput, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<SiteColumnInput. Input CSV file is available at " + SiteColumnUsageFilePath);
            Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<SiteColumnInput. Input CSV file is available at " + SiteColumnUsageFilePath);

            objSCInput = null;
            //objSCInput = ImportCsv.Read<SiteColumnInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.SiteColumnDuplicateInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objSCInput = ImportCsv.ReadMatchingColumns<SiteColumnInput>(SiteColumnUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV] [END] Read all the SiteColumns from Input and saved in List - out IEnumerable<SiteColumnInput> objSCInput, for processing.");
            Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV] [END] Read all the SiteColumns from Input and saved in List - out IEnumerable<SiteColumnInput> objSCInput, for processing.");

            try
            {
                if (objSCInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV - After Loading InputCSV");
                    bool headerPageLayout = false;
                    
                    //This is for Exception Comments:
                    exceptionCommentsInfo1 = "OLD_SiteColumn_ID: " + oldSiteColumn_ID + ", OLD_SiteColumn_InternalName: " + oldSiteColumn_InternalName + " New_SiteColumn_InternalName: " + newSiteColumn_InternalName;
                    //This is for Exception Comments:

                    //Filter - Site Column  Using SiteColumn Name Column
                    if (oldSiteColumn_InternalName != "")
                    {
                        objSCInput = from p in objSCInput
                                     where p.CustomFields_InternalName == oldSiteColumn_InternalName
                                     select p;
                    }
                    //Filter - Site Column  Using SiteColumn ID Column
                    else if (oldSiteColumn_ID != "")
                    {
                        objSCInput = from p in objSCInput
                                     where p.CustomFields_Id == oldSiteColumn_ID
                                     select p;
                    }

                    foreach (SiteColumnInput objInput in objSCInput)
                    {
                        List<SiteColumnBase> objSC_CSOMBase = SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB(outPutFolder, objInput.WebUrl, objInput.CustomFields_InternalName, objInput.CustomFields_Id, newSiteColumn_InternalName, newSiteColumn_DisplayName, Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                        
                        if (objSC_CSOMBase != null)
                        {
                            if (objSC_CSOMBase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.SiteColumnDuplicateOutput, ref objSC_CSOMBase, ref headerPageLayout);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV - After Loading InputCSV");
                    Console.WriteLine("[END] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV. Exception Message: " + ex.Message + " ExceptionComments: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "SiteColumnCSOMCreation", ex.Message, ex.ToString(), "SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ReadInputCSV");
        }
        public List<SiteColumnBase> SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB(string outPutFolder, string WebUrl, string oldSiteColumn_InternalName, string oldSiteColumn_ID, string newSiteColumn_InternalName, string newSiteColumn_DisplayName, string ActionType = Constants.ActionType_Blank, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            bool headerOutPutCSV = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<SiteColumnBase> objSC_CSOMBase = new List<SiteColumnBase>();

            ExceptionCsv.WebUrl = WebUrl;

            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                SiteColumnAndContentType_Initialization(outPutFolder, "DUPLICATE-SITE-COLUMN");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Site Column Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Site Column Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB");
                Console.WriteLine("[START] ::: SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] WebUrl is " + WebUrl);
                Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] WebUrl is " + WebUrl);
            }

            try
            {
                //Validation 
                if (Check_ValidateAllInputFields(outPutFolder, WebUrl, oldSiteColumn_InternalName, oldSiteColumn_ID, newSiteColumn_InternalName, newSiteColumn_DisplayName, UserName, Password))
                {
                    AuthenticationHelper ObjAuth = new AuthenticationHelper();
                    ClientContext clientContext = null;
                    exceptionCommentsInfo1 = "oldSiteColumn_InternalName: " + oldSiteColumn_InternalName + ", oldSiteColumn_ID: " + oldSiteColumn_ID + ", newSiteColumn_InternalName: " + newSiteColumn_InternalName + ", newSiteColumn_DisplayName: " + newSiteColumn_DisplayName;

                    //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                    if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                        clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    }
                    //SharePointOnline  => OL (Online)
                    else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                        clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    }

                    if (clientContext != null)
                    {
                        Web web = clientContext.Web;
                        FieldCollection fields = web.Fields;

                        clientContext.Load(fields);
                        clientContext.ExecuteQuery();

                        //Check if NEW SiteColumns_ISAlreadyExists
                        bool _NewSiteColumns_ISAlreadyExists = SiteColumns_ISAlreadyExists(clientContext, newSiteColumn_InternalName.ToString());
                        if (_NewSiteColumns_ISAlreadyExists)
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] This Site Column " + newSiteColumn_InternalName + " is already exist in the Web: " + WebUrl);
                            Console.WriteLine("[START][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] This Site Column " + newSiteColumn_InternalName + " is already exist in the Web: " + WebUrl);
                            return null;
                        }

                        //Get Deatils of OLD SC and Update SchemaXML and Create New SC using CSOM
                        Field _SiteColumnField = null;
                        if (oldSiteColumn_InternalName.Trim() != "")
                        {
                            _SiteColumnField = fields.GetByInternalNameOrTitle(oldSiteColumn_InternalName);
                            //Load Details
                            clientContext.Load(_SiteColumnField);
                            clientContext.ExecuteQuery();
                        }
                        else if (oldSiteColumn_ID.Trim() != "")
                        {
                            Guid oldGUID = new Guid(oldSiteColumn_ID);
                            _SiteColumnField = fields.GetById(oldGUID);
                            //Load Details
                            clientContext.Load(_SiteColumnField);
                            clientContext.ExecuteQuery();
                        }

                        if (_SiteColumnField != null)
                        {
                            //Get Updated SchemaXml
                            string _SiteColumnNewSchemaXml = SiteColumns_UpdateSiteColumnSchemaXml(_SiteColumnField, newSiteColumn_InternalName, newSiteColumn_DisplayName);

                            if (_SiteColumnNewSchemaXml != null)
                            {
                                //Create New Site Columns
                                Field _Updated_SiteColumnField = fields.AddFieldAsXml(_SiteColumnNewSchemaXml, true, AddFieldOptions.DefaultValue);
                                //Load Deatils Of Newly Created Site Column + clientContext.ExecuteQuery(); which is common for update and load command
                                clientContext.Load(_Updated_SiteColumnField);
                                clientContext.ExecuteQuery();

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Site Column is Created  " + _Updated_SiteColumnField.Id.ToString() + " and Copied Schema XML from  " + _SiteColumnField.InternalName.ToString());
                                Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Site Column is Created  " + _Updated_SiteColumnField.Id.ToString() + " and Copied Schema XML from  " + _SiteColumnField.InternalName.ToString());

                                SiteColumnBase objSCOut = new SiteColumnBase();

                                objSCOut.New_SiteColumn_ID = _Updated_SiteColumnField.Id.ToString();
                                objSCOut.New_SiteColumn_InternalName = _Updated_SiteColumnField.InternalName.ToString();
                                objSCOut.New_SiteColumn_Scope = _Updated_SiteColumnField.Scope.ToString();
                                objSCOut.New_SiteColumn_Title = _Updated_SiteColumnField.Title.ToString();
                                objSCOut.New_SiteColumn_Type = _Updated_SiteColumnField.TypeDisplayName.ToString();

                                objSCOut.Old_SiteColumn_ID = _SiteColumnField.Id.ToString();
                                objSCOut.Old_SiteColumn_InternalName = _SiteColumnField.InternalName.ToString();
                                objSCOut.old_SiteColumn_Scope = _SiteColumnField.Scope.ToString();
                                objSCOut.Old_SiteColumn_Title = _SiteColumnField.Title.ToString();
                                objSCOut.Old_SiteColumn_Type = _SiteColumnField.TypeDisplayName.ToString();

                                objSCOut.WebUrl = WebUrl;
                                objSCOut.SiteCollection = Constants.NotApplicable;
                                objSCOut.WebApplication = Constants.NotApplicable;

                                objSC_CSOMBase.Add(objSCOut);

                                //If this is called for WEB
                                if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
                                {
                                    if (objSC_CSOMBase != null)
                                    {
                                        FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.SiteColumnDuplicateOutput, ref objSC_CSOMBase,
                                            ref headerOutPutCSV);

                                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Writing the Replace Output CSV file after creating the new Site Column. Output CSV Path: " + outPutFolder + @"\" + Constants.SiteColumnDuplicateOutput);
                                        Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Writing the Replace Output CSV file after creating the new Site Column. Output CSV Path: " + outPutFolder + @"\" + Constants.SiteColumnDuplicateOutput);

                                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn for WebUrl: " + WebUrl);
                                        Console.WriteLine("[END][SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn for WebUrl: " + WebUrl);

                                        Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Site Column Trasnformation Utility Execution Completed for Web ##############");
                                        Console.WriteLine("############## Site Column Trasnformation Utility Execution Completed for Web ##############");
                                    }
                                }
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Updated Schema XML is null for New Site Column " + newSiteColumn_InternalName + " and OLD SiteColumn is " + _SiteColumnField.InternalName.ToString());
                                Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Updated Schema XML is null for New Site Column " + newSiteColumn_InternalName + " and OLD SiteColumn is " + _SiteColumnField.InternalName.ToString());
                            }
                        }
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] OLD Site Column details are null " + _SiteColumnField.InternalName.ToString());
                            Console.WriteLine("[SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] OLD Site Column details are null " + _SiteColumnField.InternalName.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "SiteColumnCSOMCreation", ex.Message, ex.ToString(), "SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            return objSC_CSOMBase;
        }
        public bool SiteColumns_ISAlreadyExists(ClientContext clientContext, string SiteColumns_InternalName)
        {
            Web web = clientContext.Web;
            clientContext.Load(web, w => w.AvailableFields);
            clientContext.ExecuteQuery();

            FieldCollection collFields = web.AvailableFields;

            // Check existing fields
            bool _IsAlreadyExists = false;

            if (SiteColumns_InternalName != "")
            {
                var fieldExists = collFields.Any(f => f.InternalName.Trim().ToLower() == SiteColumns_InternalName.Trim().ToLower());

                if (fieldExists)
                {
                    _IsAlreadyExists = true;
                    //Console.WriteLine("This Site Column " + SiteColumns_InternalName + " is already exist");
                }
            }
            else
            {
                _IsAlreadyExists = true;
                Console.WriteLine("Please enter correct Site Column Details");
            }

            return _IsAlreadyExists;
        }
        public Field SiteColumns_GetSiteColumnsDetails(ClientContext clientContext, string SiteColumns_InternalName, string SiteColumns_ID = "")
        {
            Web web = clientContext.Web;
            clientContext.Load(web, w => w.AvailableFields);
            clientContext.ExecuteQuery();

            FieldCollection collFields = web.AvailableFields;
            var field = collFields.FirstOrDefault(f => f.InternalName.Trim().ToLower() == SiteColumns_InternalName.Trim().ToLower());
            
            if (field != null)
            {         
                return field;
            }
            else
            {
                return null;
            }
        }
        public string SiteColumns_UpdateSiteColumnSchemaXml(Field _SiteColumn, string newSiteColumn_InternalName, string newSiteColumn_DisplayName)
        {
           string FieldAsXML = _SiteColumn.SchemaXml.ToString();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(FieldAsXML);
            XmlNode node = xmlDoc.SelectSingleNode("Field");
            XmlElement Xe = (XmlElement)node;

            Guid fieldID = Guid.NewGuid();

            Xe.SetAttribute("DisplayName", newSiteColumn_DisplayName);
            Xe.SetAttribute("StaticName",newSiteColumn_InternalName);
            Xe.SetAttribute("Name", newSiteColumn_InternalName);
            Xe.SetAttribute("ID", "{" + fieldID + "}");

            FieldAsXML = xmlDoc.InnerXml;

            return FieldAsXML;
        }
        #endregion

        #region AddSiteColumnToContentType
        public void AddSiteColumnToContentType_ForCSV(string ContentTypeName, string SiteColumnName, string ContentTypeUsageFilePath, string OutPutDirectory, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                SiteColumnAndContentType_Initialization(OutPutDirectory, "ADD-SITE-COLUMN-IN-CONTENT-TYPE");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add SiteColumn in ContentType Trasnformation Utility Execution Started - For InputCSV ##############");
                Console.WriteLine("############## Add SiteColumn in ContentType  Trasnformation Utility Execution Started - For InputCSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: AddSiteColumnToContentType_ForCSV");
                Console.WriteLine("[START] ::: AddSiteColumnToContentType_ForCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + OutPutDirectory);
                Console.WriteLine("[AddSiteColumnToContentType_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + OutPutDirectory);

                //Reading Input File
                IEnumerable<ContentTypeInput> objSCInput;
                AddSiteColumnToContentType_ReadInputCSV(ContentTypeName, SiteColumnName, ContentTypeUsageFilePath, OutPutDirectory, out objSCInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                Console.WriteLine("[END] ::: AddSiteColumnToContentType_ForCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: AddSiteColumnToContentType_ForCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add SiteColumn in ContentType Trasnformation Utility Execution Completed : InputCSV ##############");
                Console.WriteLine("############## Add SiteColumn in ContentType Trasnformation Utility Execution Completed : InputCSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "AddSiteColumnINContentType", ex.Message, ex.ToString(), "SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV", ex.GetType().ToString(),exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] FUNCTION SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private void AddSiteColumnToContentType_ReadInputCSV(string ContentTypeName, string SiteColumnName, string ContentTypeUsageFilePath, string outPutFolder, out IEnumerable<ContentTypeInput> objSCInput, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: AddSiteColumnToContentType_ReadInputCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ReadInputCSV] [START] Calling function ImportCsv.Read<AddSiteColumnToContentTypeInput>. Input CSV file is available at " + ContentTypeUsageFilePath);

            objSCInput = null;
            //objSCInput = ImportCsv.Read<ContentTypeInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.SiteColumnAddINContentTypeInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objSCInput = ImportCsv.ReadMatchingColumns<ContentTypeInput>(ContentTypeUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ReadInputCSV] [END] Read all the SiteColumns and ContentType from Input and saved in List - out IEnumerable<ContentTypeInput> objSCInput, for processing.");

            try
            {
                if (objSCInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] AddSiteColumnToContentType_ReadInputCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] AddSiteColumnToContentType_ReadInputCSV - After Loading InputCSV");
                    bool headerCSVColumns = false;
                   
                    //This is for Exception Comments:
                    exceptionCommentsInfo1 = "ContentTypeName: " + ContentTypeName + ", SiteColumnName: " + SiteColumnName;
                    //This is for Exception Comments:

                    //Filter - Content Type Using ContentTypeName Column
                    objSCInput = from p in objSCInput
                                 where p.ContentTypeName == ContentTypeName
                                 select p;

                    foreach (ContentTypeInput objInput in objSCInput)
                    {
                        List<AddSiteColumnToContentTypeBase> objSC_CSOMBase = AddSiteColumnToContentType_ForWeb(outPutFolder, objInput.WebUrl, objInput.ContentTypeName, SiteColumnName, Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        if (objSC_CSOMBase != null)
                        {
                            if (objSC_CSOMBase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput, ref objSC_CSOMBase, ref headerCSVColumns);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] AddSiteColumnToContentType_ReadInputCSV - After Loading InputCSV");
                    Console.WriteLine("[END] AddSiteColumnToContentType_ReadInputCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [AddSiteColumnToContentType_ReadInputCSV] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "AddSiteColumnINContentType", ex.Message, ex.ToString(), "AddSiteColumnToContentType_ReadInputCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [AddSiteColumnToContentType_ReadInputCSV] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: AddSiteColumnToContentType_ReadInputCSV");
        }
        public List<AddSiteColumnToContentTypeBase> AddSiteColumnToContentType_ForWeb(string outPutFolder, string WebUrl, string ContentTypeName, string SiteColumnName, string ActionType = "", string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            bool headerCSVColumns = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<AddSiteColumnToContentTypeBase> objSC_CSOMBase = new List<AddSiteColumnToContentTypeBase>();

            ExceptionCsv.WebUrl = WebUrl;

            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                SiteColumnAndContentType_Initialization(outPutFolder, "ADD-SITE-COLUMN-IN-CONTENT-TYPE");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add SiteColumn in ContentType Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Add SiteColumn in ContentType Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: AddSiteColumnToContentType_ForWeb");
                Console.WriteLine("[START] ::: AddSiteColumnToContentType_ForWeb");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[AddSiteColumnToContentType_ForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb] WebUrl is " + WebUrl);
                Console.WriteLine("[AddSiteColumnToContentType_ForWeb] WebUrl is " + WebUrl);
            }

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                exceptionCommentsInfo1 = "ContentTypeName: " + ContentTypeName + ", SiteColumnName: " + SiteColumnName + ", WebUrl: " + WebUrl;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][AddSiteColumnToContentType_ForWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][AddSiteColumnToContentType_ForWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][AddSiteColumnToContentType_ForWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][AddSiteColumnToContentType_ForWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;

                    // Try to load the new content type
                    var contentType = GetContentTypeByName(clientContext, web, ContentTypeName);
                    if (contentType == null)
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb] Content Type " + ContentTypeName + " does not exists in this Web: " + WebUrl + " OR Content Type Internal Name is in correct");
                        Console.WriteLine("[AddSiteColumnToContentType_ForWeb] Content Type " + ContentTypeName + " does not exists in this Web: " + WebUrl + " OR Content Type Internal Name is in correct");
                        return null; // not found
                    }

                    // Load field links to content type
                    clientContext.Load(contentType);
                    clientContext.Load(contentType.FieldLinks);
                    clientContext.ExecuteQuery();

                    // Try to load the new field
                    Field fld = null;
                    fld = web.Fields.GetByInternalNameOrTitle(SiteColumnName);
                    clientContext.Load(fld);
                    clientContext.ExecuteQuery();

                    if(fld != null)
                    {
                        // Try to load the content type/site column connection
                        var hasFieldConnected = contentType.FieldLinks.Any(f => f.Name == SiteColumnName);

                        // Reference exists already, no further action required
                        if (hasFieldConnected)
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb]\"" + SiteColumnName + "\": This Site Column is already present in this Content Type " + ContentTypeName);
                            Console.WriteLine("[AddSiteColumnToContentType_ForWeb]\"" + SiteColumnName + "\": This Site Column is already present in this Content Type " + ContentTypeName);
                            return null;
                        }

                        // Reference does not exist yet - create the connection
                        FieldLinkCreationInformation link = new FieldLinkCreationInformation();
                        link.Field = fld;
                        contentType.FieldLinks.Add(link);

                        contentType.Update(true);
                        clientContext.ExecuteQuery();

                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb][Site Column \"" + SiteColumnName + "\" is added in Content Type " + ContentTypeName);
                        Console.WriteLine("[AddSiteColumnToContentType_ForWeb][Site Column \"" + SiteColumnName + "\" is added in Content Type " + ContentTypeName);

                        AddSiteColumnToContentTypeBase objSCOut = new AddSiteColumnToContentTypeBase();
                        
                        objSCOut.ContentTypeName = contentType.Name;
                        objSCOut.SiteColumnName = fld.InternalName;
                        objSCOut.ContentTypeID = contentType.Id.ToString();
                        objSCOut.SiteColumnID = fld.Id.ToString();

                        objSCOut.WebUrl = WebUrl;
                        objSCOut.SiteCollection = Constants.NotApplicable;
                        objSCOut.WebApplication = Constants.NotApplicable;

                        objSC_CSOMBase.Add(objSCOut);

                        //If this for WEB
                        if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
                        {
                            if (objSC_CSOMBase != null)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput, ref objSC_CSOMBase,
                                    ref headerCSVColumns);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType_ForWeb] Writing the Replace Output CSV file after adding Site Column in Content Type. Output CSV Path: " + outPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput);
                                Console.WriteLine("[AddSiteColumnToContentType_ForWeb]  Writing the Replace Output CSV file after adding Site Column in Content Type. Output CSV Path: " + outPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][AddSiteColumnToContentType_ForWeb] AddSiteColumnToContentType_ForWeb for WebUrl: " + WebUrl);
                                Console.WriteLine("[END][AddSiteColumnToContentType_ForWeb] AddSiteColumnToContentType_ForWeb for WebUrl: " + WebUrl);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add SiteColumn in ContentType Trasnformation Utility Execution Completed for Web ##############");
                                Console.WriteLine("############## Add SiteColumn in ContentType Trasnformation Utility Execution Completed for Web ##############");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "AddSiteColumnINContentType", ex.Message, ex.ToString(), "AddSiteColumnToContentType_ForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [AddSiteColumnToContentType_ForWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [AddSiteColumnToContentType_ForWeb] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            return objSC_CSOMBase;
        }
        public void AddSiteColumnToContentType(ClientContext clientContext, ContentType newContentType, ContentType oldContentType)
        {
            string exceptionCommentsInfo1 = string.Empty;
            
            try
            {
                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    
                    // Load field links to content type
                    clientContext.Load(newContentType);
                    clientContext.Load(newContentType.FieldLinks);
                    clientContext.ExecuteQuery();
                    
                    //Exception Comments
                    exceptionCommentsInfo1 = "New ContentTypeName: " + newContentType.Name + "Old ContentTypeName: " + oldContentType.Name + "";

                    //Load all fields of old ContentType
                    FieldCollection oldContentTypesFieldColl = oldContentType.Fields;
                    clientContext.Load(oldContentTypesFieldColl);
                    clientContext.ExecuteQuery();
                    
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType][START]Adding Site Columns in New ContentTypeName: " + newContentType.Name + " from Old ContentTypeName: " + oldContentType.Name);
                    Console.WriteLine("[AddSiteColumnToContentType][START]Adding Site Columns in New ContentTypeName: " + newContentType.Name + " from Old ContentTypeName: " + oldContentType.Name);
                    
                    foreach (Field field in oldContentTypesFieldColl)
                    {
                        clientContext.Load(field);
                        clientContext.ExecuteQuery();
                        // Try to load the content type/site column connection
                        var hasFieldConnected = newContentType.FieldLinks.Any(f => f.Name == field.InternalName);

                        // Reference exists already, do not add again it, in Content Type
                        if (hasFieldConnected)
                        {
                            ///As per the suggestion, Commented oot the below message:

                            //Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType]\"" + field.InternalName + "\": This Site Column is already present in this Content Type " + newContentType.Name);
                            //Console.WriteLine("[AddSiteColumnToContentType]\"" + field.InternalName + "\": This Site Column is already present in this Content Type " + newContentType.Name);
                        }
                        else
                        {
                            // Reference does not exist yet - create the connection and Add Site Column in this Content Type
                            FieldLinkCreationInformation link = new FieldLinkCreationInformation();
                            link.Field = field;
                            newContentType.FieldLinks.Add(link);
                            newContentType.Update(true);
                            clientContext.ExecuteQuery();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType]Site Column \"" + field.InternalName + "\" is added successfully in Content Type " + newContentType.Name);
                            Console.WriteLine("[AddSiteColumnToContentType]Site Column \"" + field.InternalName + "\" is added successfully in Content Type " + newContentType.Name);
                        }
                    }
                    
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddSiteColumnToContentType][END] Added Site Columns in New ContentTypeName: " + newContentType.Name + " from Old ContentTypeName: " + oldContentType.Name);
                    Console.WriteLine("[AddSiteColumnToContentType][END] Added Site Columns in New ContentTypeName: " + newContentType.Name + " from Old ContentTypeName: " + oldContentType.Name);   
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "AddSiteColumnINContentType", ex.Message, ex.ToString(), "AddSiteColumnToContentType", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [AddSiteColumnToContentType] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [AddSiteColumnToContentType] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        #endregion

        #region ReplaceContentTypeInList
        public void ReplaceContentTypeinList_ForCSV(string DiscoveryUsage_OutPutFolder, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                SiteColumnAndContentType_Initialization(DiscoveryUsage_OutPutFolder, "REPLACE-CONTENT-TYPE-IN-LIST");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Update ContentType in List Trasnformation Utility Execution Started - For InputCSV ##############");
                Console.WriteLine("############## Add/Update ContentType in List Trasnformation Utility Execution Started - For InputCSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ReplaceContentTypeinList_ForCSV");
                Console.WriteLine("[START] ::: ReplaceContentTypeinList_ForCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + DiscoveryUsage_OutPutFolder);
                Console.WriteLine("[ReplaceContentTypeinList_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + DiscoveryUsage_OutPutFolder);

                //Reading Input File
                IEnumerable<UpdateContentTypeinListInput> objListCTInput;
                ReplaceContentTypeinList_ReadInputCSV(DiscoveryUsage_OutPutFolder, out objListCTInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                Console.WriteLine("[END] ::: ReplaceContentTypeinList_ForCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ReplaceContentTypeinList_ForCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Update ContentType in List Trasnformation Utility Execution Completed : InputCSV ##############");
                Console.WriteLine("############## Add/Update ContentType in List Trasnformation Utility Execution Completed : InputCSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ReplaceContentTypeinList_ForCSV. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "REPLACE-CONTENT-TYPE-IN-LIST", ex.Message, ex.ToString(), "ReplaceContentTypeinList_ForCSV", ex.GetType().ToString(),exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] FUNCTION ReplaceContentTypeinList_ForCSV. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private void ReplaceContentTypeinList_ReadInputCSV(string outPutFolder, out IEnumerable<UpdateContentTypeinListInput> objListCTInput, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ReplaceContentTypeinList_ReadInputCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ReadInputCSV] [START] Calling function ImportCsv.Read<UpdateContentTypeinListInput>. Input CSV file is available at " + outPutFolder + " and Input file name is " + Constants.ContentType_Add_To_ListInput);

            objListCTInput = null;
            objListCTInput = ImportCsv.Read<UpdateContentTypeinListInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.ContentType_Add_To_ListInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ReadInputCSV] [END] Read all the Lists and ContentType from Input and saved in List - out IEnumerable<AddSiteColumnToContentTypeInput> , for processing.");

            try
            {
                if (objListCTInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ReplaceContentTypeinList_ReadInputCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] ReplaceContentTypeinList_ReadInputCSV - After Loading InputCSV");
                    bool headerPageLayout = false;

                    foreach (UpdateContentTypeinListInput objInput in objListCTInput)
                    {
                        //This is for Exception Comments:
                        exceptionCommentsInfo1 = "newContentTypeName: " + objInput.newContentTypeName + ", WebUrl: " + objInput.WebUrl + ", oldContentTypeId: " + objInput.oldContentTypeId;
                        //This is for Exception Comments:

                        List<UpdateContentTypeinListBase> objListCT_CSOMBase = ReplaceContentTypeinList_ForWeb(outPutFolder, objInput.WebUrl, objInput.ListName, objInput.oldContentTypeId, objInput.newContentTypeName, "CSVUpdates", SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        if (objListCT_CSOMBase != null)
                        {
                            if (objListCT_CSOMBase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.SiteColumnAddINContentTypeOutput, ref objListCT_CSOMBase, ref headerPageLayout);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ReplaceContentTypeinList_ReadInputCSV - After Loading InputCSV");
                    Console.WriteLine("[END] ReplaceContentTypeinList_ReadInputCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ReplaceContentTypeinList_ReadInputCSV. Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "REPLACE-CONTENT-TYPE-IN-LIST", ex.Message, ex.ToString(), "ReplaceContentTypeinList_ReadInputCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ReplaceContentTypeinList_ReadInputCSV] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ReplaceContentTypeinList_ReadInputCSV");
        }
        public List<UpdateContentTypeinListBase> ReplaceContentTypeinList_ForWeb(string outPutFolder, string WebUrl, string ListName, string oldContentTypeId, string newContentTypeName, string ActionType = "", string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            bool headerMasterPage = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<UpdateContentTypeinListBase> objList_CTBase = new List<UpdateContentTypeinListBase>();

            ExceptionCsv.WebUrl = WebUrl;

            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                SiteColumnAndContentType_Initialization(outPutFolder, "REPLACE-CONTENT-TYPE-IN-LIST");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Update ContentType in List Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Add/Update ContentType in List Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ReplaceContentTypeinList_ForWeb");
                Console.WriteLine("[START] ::: ReplaceContentTypeinList_ForWeb");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ReplaceContentTypeinList_ForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] WebUrl is " + WebUrl);
                Console.WriteLine("[ReplaceContentTypeinList_ForWeb] WebUrl is " + WebUrl);
            }

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                exceptionCommentsInfo1 = "ListName: " + ListName + ", oldContentTypeId: " + oldContentTypeId + ", newContentTypeName: " + newContentTypeName;


                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ReplaceContentTypeinList_ForWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ReplaceContentTypeinList_ForWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ReplaceContentTypeinList_ForWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ReplaceContentTypeinList_ForWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;

                    // Get content type and list
                    ContentType newContentType = GetContentTypeByName(clientContext, web, newContentTypeName);
                    if (newContentType == null)
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] New Content Type " + newContentTypeName + " does not exists in this Web: " + WebUrl + " OR Content Type Internal Name is in correct");
                        Console.WriteLine("[ReplaceContentTypeinList_ForWeb] New Content Type " + newContentTypeName + " does not exists in this Web: " + WebUrl + " OR Content Type Internal Name is in correct");
                        return null; // not found
                    }

                    ListCollection lists = web.Lists;
                    // Load all data required

                    clientContext.Load(newContentType);

                    clientContext.Load(lists,
                            l => l.Include(list => list.ContentTypes));
                    clientContext.Load(lists);
                    clientContext.ExecuteQuery();

                    var listsWithContentType = new List<List>();

                    foreach (List list in lists)
                    {
                        //If User Pass List Name
                        if (ListName != "")
                        {
                            if (list.Title.ToString().Trim().ToLower() == ListName.ToString().Trim().ToLower())
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] Processing for List: " + ListName);
                                Console.WriteLine("[ReplaceContentTypeinList_ForWeb] Processing for List: " + ListName);

                                bool hasOldContentType = list.ContentTypes.Any(c => c.StringId.StartsWith(oldContentTypeId));
                                if (hasOldContentType)
                                {
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] This list has old content type " + oldContentTypeId);
                                    Console.WriteLine("[ReplaceContentTypeinList_ForWeb] This list has old content type " + oldContentTypeId);

                                    listsWithContentType.Add(list);
                                    break;
                                }
                            }
                        }
                        //Else Execute for all Lists => web.Lists;
                        else
                        {
                            bool hasOldContentType = list.ContentTypes.Any(c => c.StringId.StartsWith(oldContentTypeId));
                            if (hasOldContentType)
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] The list " + list.Title.ToString() + " has old content type " + oldContentTypeId);
                                Console.WriteLine("[ReplaceContentTypeinList_ForWeb] The list " + list.Title.ToString() + " has old content type " + oldContentTypeId);

                                listsWithContentType.Add(list);
                            }
                        }
                    }

                    foreach (List list in listsWithContentType)
                    {
                        // Check if the new content type is already attached to the library
                        var listHasContentTypeAttached = list.ContentTypes.Any(c => c.Name == newContentTypeName);
                        if (!listHasContentTypeAttached)
                        {
                            // Attach content type to list
                            list.ContentTypes.AddExistingContentType(newContentType);
                            clientContext.ExecuteQuery();
                        }


                        // Load all list items
                        CamlQuery query = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection items = list.GetItems(query);
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();
                        int _itemsCount = 0;
                        _itemsCount = items.Count;

                        if (_itemsCount > 0)
                        {
                            // For each list item check if it is set to the old content type, update to new one if required
                            foreach (ListItem listItem in items)
                            {
                                // Currently assigned content type to this item
                                var currentContentTypeId = listItem["ContentTypeId"] + "";
                                var isOldContentTypeAssigned = currentContentTypeId.StartsWith(oldContentTypeId);
                                
                                // This item is not assigned to the old content type - skip to next one
                                if (!isOldContentTypeAssigned) continue;

                                // Update to new content type
                                listItem["ContentTypeId"] = newContentType.StringId; // newContentTypeId;
                                
                                listItem.Update();
                            }
                        }
                        // Submit all changes
                        clientContext.ExecuteQuery();

                        //If Content Type is attched
                        if (!listHasContentTypeAttached)
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb][Added New Content Type " + newContentType.StringId + ", where old Content Type was " + oldContentTypeId);
                            Console.WriteLine("[ReplaceContentTypeinList_ForWeb][Added New Content Type " + newContentType.StringId + ", where old Content Type was " + oldContentTypeId);

                            UpdateContentTypeinListBase objListCTOut = new UpdateContentTypeinListBase();

                            objListCTOut.ListName = list.Title.ToString();
                            objListCTOut.newContentTypeName = newContentType.StringId;
                            objListCTOut.oldContentTypeId = oldContentTypeId.ToString();

                            objListCTOut.WebUrl = WebUrl;
                            objListCTOut.SiteCollection = Constants.NotApplicable;
                            objListCTOut.WebApplication = Constants.NotApplicable;

                            objList_CTBase.Add(objListCTOut);
                        }
                    }

                    //If this for WEB
                    if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
                    {
                        if (objList_CTBase != null)
                        {
                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ContentType_Add_To_ListOutput, ref objList_CTBase,
                                ref headerMasterPage);

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceContentTypeinList_ForWeb] Writing the Replace Output CSV file after adding newcontent type in list - FileUtility.WriteCsVintoFile");
                            Console.WriteLine("[ReplaceContentTypeinList_ForWeb]  Writing the Replace Output CSV file after adding newcontent type in list  ");

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ReplaceContentTypeinList_ForWeb] ReplaceContentTypeinList_ForWeb for WebUrl: " + WebUrl);
                            Console.WriteLine("[END][ReplaceContentTypeinList_ForWeb] ReplaceContentTypeinList_ForWeb for WebUrl: " + WebUrl);

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Update ContentType in List Trasnformation Utility Execution Completed for Web ##############");
                            Console.WriteLine("############## Add/Update ContentType in List Trasnformation Utility Execution Completed for Web ##############");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "REPLACE-CONTENT-TYPE-IN-LIST", ex.Message, ex.ToString(), "ReplaceContentTypeinList_ForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ReplaceContentTypeinList_ForWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ReplaceContentTypeinList_ForWeb] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            return objList_CTBase;

        }
        #endregion

        #region ContentType Creation
        public void ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV(string oldContentTypeName, string newContentTypeName, string ContentTypeUsageFilePath, string OutPutFolder, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                SiteColumnAndContentType_Initialization(OutPutFolder, "DUPLICATE-CONTENT-TYPE");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Add/Duplicate ContentType Trasnformation Utility Execution Started - Using ContentType CSV ##############");
                Console.WriteLine("############## Add/Duplicate ContentType Trasnformation Utility Execution Started - Using ContentType CSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV");
                Console.WriteLine("[START] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in folder/path" + OutPutFolder);
                Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in folder/path" + OutPutFolder);

                //Reading Input File
                IEnumerable<ContentTypeInput> objCTInput;
                ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV(oldContentTypeName, newContentTypeName, ContentTypeUsageFilePath, OutPutFolder, out objCTInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                Console.WriteLine("[END] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Duplicate ContentType Trasnformation Utility Execution Completed - Using ContentType CSV  ##############");
                Console.WriteLine("############## Add/Duplicate ContentType Trasnformation Utility Execution Completed - Using ContentType CSV  ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV. Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "DUPLICATE-CONTENT-TYPE", ex.Message, ex.ToString(), "ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV", ex.GetType().ToString(),exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] FUNCTION ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV. Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private void ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV(string oldContentTypeName, string newContentTypeName, string ContentTypeUsageFilePath, string outPutFolder, out IEnumerable<ContentTypeInput> objCTInput, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<ContentTypeInput>. Input CSV file is available at " + ContentTypeUsageFilePath);
            Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<ContentTypeInput>. Input CSV file is available at " + ContentTypeUsageFilePath);
            
            objCTInput = null;
            //objCTInput = ImportCsv.Read<ContentTypeInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.ContentTypeDuplicateInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objCTInput = ImportCsv.ReadMatchingColumns<ContentTypeInput>(ContentTypeUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] [END] Read all the Lists and ContentType from Input and saved in List - out IEnumerable<ContentTypeInput> objCTInput , for processing.");
            Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] [END] Read all the Lists and ContentType from Input and saved in List - out IEnumerable<ContentTypeInput> objCTInput , for processing.");
            
            try
            {
                if (objCTInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV - After Loading InputCSV");
                    bool headerPageLayout = false;

                    //Filter - Content Type Using ContentTypeName Column
                    objCTInput = from p in objCTInput
                                 where p.ContentTypeName == oldContentTypeName
                                  select p;
                    
                    //This is for Exception Comments:
                    exceptionCommentsInfo1 = "NewContentTypeName: " + newContentTypeName + ", OldContentTypeName: " + oldContentTypeName;
                    //This is for Exception Comments:

                    foreach (ContentTypeInput objInput in objCTInput)
                    {
                        List<ContentTypeBase> objListCT_CSOMBase = ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB(outPutFolder, objInput.WebUrl, objInput.ContentTypeName, newContentTypeName, Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        if (objListCT_CSOMBase != null)
                        {
                            if (objListCT_CSOMBase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ContentTypeDuplicateOutput, ref objListCT_CSOMBase, ref headerPageLayout);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV - After Loading InputCSV");
                    Console.WriteLine("[END] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "DUPLICATE-CONTENT-TYPE", ex.Message, ex.ToString(), "ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ReadInputCSV");
        }
        public List<ContentTypeBase> ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB(string outPutFolder, string WebUrl, string oldContentTypeName, string newContentTypeName, string ActionType = "", string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            bool headerCSVColumns = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<ContentTypeBase> objCTBase = new List<ContentTypeBase>();

            ExceptionCsv.WebUrl = WebUrl;

            if (ActionType.ToLower().Trim() == Constants.ActionType_Web.ToLower())
            {
                SiteColumnAndContentType_Initialization(outPutFolder, "DUPLICATE-CONTENT-TYPE");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Duplicate ContentType Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Add/Duplicate ContentType Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB");
                Console.WriteLine("[START] ::: ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] WebUrl is " + WebUrl);
                Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] WebUrl is " + WebUrl);
            }

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                exceptionCommentsInfo1 = "oldContentTypeName: " + oldContentTypeName + ", newContentTypeName: " + newContentTypeName;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;

                    if(oldContentTypeName.ToString() != "")
                    {
                        // Content type is present or not ??
                        var OLD_ContentType = GetContentTypeByName(clientContext, web, oldContentTypeName);                    
                        if (OLD_ContentType != null)
                        {
                            //We found the old content type in this Context
                            clientContext.Load(OLD_ContentType);
                            clientContext.ExecuteQuery();

                            // Check if the New content type does not exist yet
                            var NEW_ContentType = GetContentTypeByName(clientContext, web, newContentTypeName);
                            // Content type exists already, no further action required
                            if (NEW_ContentType != null)
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] New Content Type " + newContentTypeName + " is already exists in this Web: " + WebUrl);
                                Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] New Content Type " + newContentTypeName + " is already exists in this Web: " + WebUrl);
                                return null;
                            }

                            // Create a Content Type Information object
                            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
                            newCt.Name = newContentTypeName;
                            newCt.Description = OLD_ContentType.Description;
                            newCt.Group = OLD_ContentType.Group;
                            newCt.ParentContentType = OLD_ContentType.Parent;

                            ContentType myContentType = web.ContentTypes.Add(newCt);
                            clientContext.ExecuteQuery();

                            //Load newlyCreatedContentType, to write in OUTPUT CSV
                            var newlyCreatedContentType = GetContentTypeByName(clientContext, web, newContentTypeName);

                            if (newlyCreatedContentType != null)
                            {
                                clientContext.Load(newlyCreatedContentType);
                                clientContext.ExecuteQuery();

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] ContentTypeName:" + newlyCreatedContentType.Name + ", ID: " + newlyCreatedContentType.Id.ToString() + "  created successfully in Web: " + WebUrl);
                                Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB]  ContentTypeName:" + newlyCreatedContentType.Name + ", ID: " + newlyCreatedContentType.Id.ToString() + " created successfully in Web: " + WebUrl);

                                /// Add Site Columns in Newly Created ContentType
                                AddSiteColumnToContentType(clientContext, newlyCreatedContentType, OLD_ContentType);
                                /// Add Site Columns in Newly Created ContentType

                                
                                ContentTypeBase objCT = new ContentTypeBase();
                                objCT.OldContentTypeName = oldContentTypeName;
                                objCT.NewContentTypeName = newContentTypeName;
                                objCT.NewContentTypeID = newlyCreatedContentType.Id.ToString();
                                objCT.OldContentTypeID = OLD_ContentType.Id.ToString();

                                objCT.WebUrl = WebUrl;
                                objCT.SiteCollection = Constants.NotApplicable;
                                objCT.WebApplication = Constants.NotApplicable;

                                objCTBase.Add(objCT);
                            }

                            //If ==> This is for WEB
                            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
                            {
                                if (objCTBase != null)
                                {
                                    FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ContentTypeDuplicateOutput, ref objCTBase,
                                        ref headerCSVColumns);

                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Writing the Replace Output CSV file after creating a new content type. Output CSV Path: " + outPutFolder + @"\" + Constants.ContentTypeDuplicateOutput);
                                    Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB]  Writing the Replace Output CSV file after creating a new content type. Output CSV Path: " + outPutFolder + @"\" + Constants.ContentTypeDuplicateOutput);

                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB for WebUrl: " + WebUrl);
                                    Console.WriteLine("[END][ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB for WebUrl: " + WebUrl);

                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Add/Duplicate ContentType Trasnformation Utility Execution Completed for Web ##############");
                                    Console.WriteLine("############## Add/Duplicate ContentType Trasnformation Utility Execution Completed for Web ##############");
                                }
                            }
                            //If ==> This is for WEB
                        }
                        else
                        {
                            //Old Content Type is not present in this Context
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Old Content Type " +oldContentTypeName + " does not exists in this Web: " + WebUrl);
                            Console.WriteLine("[ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Old Content Type " + oldContentTypeName + " does not exists in this Web: " + WebUrl);  
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "DUPLICATE-CONTENT-TYPE", ex.Message, ex.ToString(), "ReplaceContentTypeinList_ForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            return objCTBase;
        }
        public FieldCollection ContentType_GetAssociatedSiteColumnsForContentType(ClientContext clientContext, string ContentTypeName, string ContentTypeID)
        {
            ContentType _ContentType = null;
            FieldCollection CT_fieldColl = null;

            if(ContentTypeName!= "")
            {
                _ContentType = GetContentTypeByName(clientContext, clientContext.Web, ContentTypeName, false);
            }
            else if (ContentTypeID != "")
            {
                _ContentType = GetContentTypeByID(clientContext, clientContext.Web, ContentTypeID, false);
            }

            /// Gets a value that specifies the collection of fields for the content type
            CT_fieldColl = _ContentType.Fields;

            clientContext.Load(CT_fieldColl);
            clientContext.ExecuteQuery();
            
            /*
            // Display the field name
            foreach (Field field in fieldColl)
            {
                Console.WriteLine(field.Title);
                Console.WriteLine(field.Id);
            }
            */

            return CT_fieldColl;
        }
        #endregion
    }
}
