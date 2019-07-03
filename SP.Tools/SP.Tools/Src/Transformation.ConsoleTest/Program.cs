using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.MasterPage;
using Transformation.PowerShell.PageLayouts;
using Transformation.PowerShell.WebPart;
using Transformation.PowerShell.SiteColumnAndContentTypes;
using Microsoft.SharePoint.Client;

namespace Transformation.ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime startTime = DateTime.Now;
            Console.WriteLine("Start Time: " + startTime);

            if((ConfigurationManager.AppSettings["operation"]).ToLower() == "get")
                GetWebPartUsage();
            else if ((ConfigurationManager.AppSettings["operation"]).ToLower() == "replace")
                ReplaceWebPartByCSV();
            else if ((ConfigurationManager.AppSettings["operation"]).ToLower() == "delete")
                DeleteWebPart_UsingCSV();


         //GetWebPartUsage();
            //GetWebPartProperties();
            //ConfigureNewWebPartXml();
            //DeleteWebPart();
            //AddWebPart();
            //GetWebPartProperties_UsingCSV();
            //ConfigureNewWebPartXmlBulk();
           
            //AddWebPart_UsingCSV();
            //ReplaceWebPart();
           //ReplaceWebPartByCSV();
            //UploadDependencyFile();
            //TransformWebPart_UsingCSV();

           // CT_ContentTypesCommands();
            //MasterPage_Trasnformation();
            Console.WriteLine("Finish");

            DateTime endTime = DateTime.Now;
            TimeSpan ts = endTime - startTime;


            Console.WriteLine("End Time: " + endTime);
            Console.WriteLine("Time Span: " + ts.TotalSeconds);

            Console.ReadLine();

        }


        public static void CT_ContentTypesCommands()
        {
            //228
            String OutPutDirectory = @"E:\VirendraKumar\ProjectTest\ContentType";
            string WebUrl = "http://001d-cam-tap01:8899/sites/TestContosoPublishingSite";
            string oldContentTypeName = "ContosoLibraryContentType";
            string newContentTypeName = "ContosoLibraryContentTypeCSV6";

            SiteColumnAndContentTypeHelper objSC = new SiteColumnAndContentTypeHelper();
            //objSC.ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB(OutPutDirectory,WebUrl, oldContentTypeName, newContentTypeName, "web", "OP", "ms-mla-paraja", "Password123", "mgmt7");
            //objSC.ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV(oldContentTypeName, newContentTypeName, @"E:\VirendraKumar\ProjectTest\ContentType_Usage.csv", OutPutDirectory, "OP", "ms-mla-paraja", "Password123", "mgmt7");

            //objSC.AddSiteColumnToContentType_ForWeb(@"E:\VirendraKumar\ProjectTest\SiteColumnsContentType", "https://intranet.poc.com/sites/TestContosoPublishingSite/", "ContosoLibraryContentTypeCSV6", "ContosoStatusNew", "web", "OP", "ms-mla-paraja", "Password123", "mgmt7");
            objSC.AddSiteColumnToContentType_ForCSV("ContosoLibraryContentType", "ContosoStatusNew", @"E:\VirendraKumar\ProjectTest\ContentType_Usage.csv", @"E:\VirendraKumar\ProjectTest\SiteColumnsContentType", "OP", "ms-mla-paraja", "Password123", "mgmt7");


            //objSC.ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForWEB(@"E:\VirendraKumar\ProjectTest\SiteColumnsContentType", "https://intranet.campoc.com/sites/T_Master_Page_Offshore/", "Contoso Document", "NewCT23Apr3", "web", "OP", "ms-mla-paraja", "Password123", "mgmt7");
            //objSC.ReplaceContentTypeinList_ForWeb(@"E:\VirendraKumar\ProjectTest\SiteColumnsContentType", "https://intranet.campoc.com/sites/T_Master_Page_Offshore/", "VK_C_List ", "0x010090E84F7F337045DB9DCB9F4D3688DB8B00E05E98E692502F4AA55DB54546B2639F", "NewCT22Apr", "web", "OP", "ms-mla-paraja", "Password123", "mgmt7");
            //objSC.SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_ForWEB(@"E:\VirendraKumar\ProjectTest\SiteColumnsContentType", "https://intranet.campoc.com/sites/T_Master_Page_Offshore/", "Contoso Status", "8478039d-fbd5-421d-bd6c-87a07d7ce499", "SC22Apr", "SC22Apr", "web", "OP", "ms-mla-paraja", "Password123", "mgmt7");

            //objSC.SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV(@"E:\VirendraKumar\ProjectTest", "OP", "ms-mla-paraja", "Password123", "mgmt7");
            /*using (var cc = new ClientContext("https://intranet.campoc.com/sites/T_Master_Page_Offshore/"))
            {
                //https://intranet.campoc.com/sites/TestListWorkFlowAssociationsContoso

                //https://intranet.campoc.com/sites/T_Master_Page_Offshore/


                string internalName = "ContosoStatusxx";
                string oldSiteColumn_ID = "8478039d-fbd5-421d-bd6c-87a07d7ce499";
                FieldCollection fields = cc.Web.Fields;
                //cc.Web.Context.Load(fields, fc => fc.Include(f => f.Id, f => f.InternalName));
                //cc.Web.Context.ExecuteQuery();
                
                //bool yes= SiteColumns_ISAlreadyExists(cc, internalName);
                //Console.WriteLine(yes);
                //Field field = SiteColumns_GetSiteColumnsDetails(cc, internalName);
                //var field = fields.FirstOrDefault(f => f.InternalName == internalName);
                //Field field = fields.GetByInternalNameOrTitle(internalName);
                Guid oldGUID = new Guid(oldSiteColumn_ID);
                Field field = fields.GetById(oldGUID);

               
                try
                {
                    if (field != null && field.ToString() != "")
                    {
                        cc.Web.Context.Load(field);
                        cc.Web.Context.ExecuteQuery();

                        Console.WriteLine(field.Id.ToString());
                        Console.WriteLine(field.InternalName.ToString());
                        Console.WriteLine(field.Title.ToString());
                        Console.WriteLine(field.Hidden.ToString());
                        Console.WriteLine(field.CanBeDeleted.ToString());
                        Console.WriteLine(field.SchemaXml.ToString());
                        //Console.WriteLine("YESSSSSSSSSS");
                    }
                }
                catch (Exception ex)
                { Console.WriteLine(ex.Message.ToString()); }

                //SiteColumns_ISAlreadyExists(cc, internalName, "");
                
                //Field field = SiteColumns_GetSiteColumnsDetails(cc, internalName, "");
                //Console.WriteLine(field.Id.ToString());
                
                //Console.WriteLine(field.InternalName.ToString());
                //Console.WriteLine(field.Title.ToString());

                //Console.WriteLine(field.Hidden.ToString());
                //Console.WriteLine(field.CanBeDeleted.ToString());
                //Console.WriteLine(field.SchemaXml.ToString());
                //ContentType CT = GetContentTypeByName(cc, cc.Web, "ContosoLibraryContentTypeNEWWW2");
                //ContosoLibraryContentTypeNEWWW2
                //Console.WriteLine(CT.SchemaXml.ToString());
            }*/
        }
        public static void MasterPage_Trasnformation()
        {
            MasterPageHelper objMasterHelper = new MasterPageHelper();
            objMasterHelper.ChangeMasterPageForWeb(@"E:\VirendraKumar\ProjectTest\MasterPage", "https://intranet.poc.com/sites/TestContosoPublishingSite", "contoso.master", "seattle.master", true, true, "", "OP", "ms-mla-paraja", "Password123", "mgmt7");

        }
        public static void GetSiteColumnsListForContentType()
        {
            //// String Variable to store the siteURL
            string siteURL = "https://intranet.campoc.com/sites/T_Master_Page_Offshore/";

            //// Get the context for the SharePoint Site to access the data
            ClientContext clientContext = new ClientContext(siteURL);

            /*//// Get the content type using ID: 0x01003D7B5A54BF843D4381F54AB9D229F98A - is the ID of the "Custom" content Type
            ContentType ct = clientContext.Web.ContentTypes.GetById("0x01010002CB55784047481FA81FC315856D4A41");

            //// Gets a value that specifies the collection of fields for the content type
            FieldCollection fieldColl = ct.Fields;

            clientContext.Load(fieldColl);
            clientContext.ExecuteQuery();

            //// Display the field name
            foreach (Field field in fieldColl)
            {
                Console.WriteLine(field.Title);
                Console.WriteLine(field.Id);
            }*/

            SiteColumnAndContentTypeHelper objSC = new SiteColumnAndContentTypeHelper();
            ContentType ct = objSC.GetContentTypeByName(clientContext, "Contoso Document");
            //// Gets a value that specifies the collection of fields for the content type
            FieldCollection fieldColl = ct.Fields;

            clientContext.Load(fieldColl);
            clientContext.ExecuteQuery();

            //// Display the field name
            foreach (Field field in fieldColl)
            {
                Console.WriteLine(field.Title);
                Console.WriteLine(field.Id);
            }

        }

        public static void TransformWebPart_UsingCSV()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            //string WebUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";
            //string WebPartType = "Content Editor";
            // string WebPartType = "RemoteContentQueryWebPart";
            string WebPartType = "TeaserWebPart";
            //string WebPartType = "WelcomeWebPart";
            //String OutPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\TransformWebPartAcrossWeb";
            String OutPutDirectory = @"E:\VirendraKumar\ProjectTest\TransformWebPartAcrossWeb";
            string targetWebPartFileName = "N/A";
            string targetWebPartXmlFile = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\AppPartXml.xml";
            //string UsageFile = @"C:\DiscoveryTool\CAM\WebParts_Usage.csv";
            string UsageFile = @"E:\VirendraKumar\ProjectTest\WebParts_Usage.csv";
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.TransformWebPart_UsingCSV(UsageFile, WebPartType, targetWebPartFileName, targetWebPartXmlFile, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
        
        public static void UploadDependencyFile()
        {
            string folderServerRelativeUrl = "/sites/TestContosoPublishingSite/_catalogs/masterpage/Display Templates/Content Web Parts";
            string fileName = "Item_PictureOnTop_MyCustom.html";
            string localFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\UploadDependencyFile\Item_PictureOnTop_MyCustom.html";
            bool overwriteIfExists = true;
            string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\UploadDependencyFile";
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string webUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.UploadDependencyFile(webUrl, folderServerRelativeUrl, fileName, localFilePath, overwriteIfExists, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void GetWebPartUsage()
        {
            //string SharePointOnline_OR_OnPremise = "OP";
            //string UserName =ConfigurationManager.AppSettings["userName"];
            //string Password = ConfigurationManager.AppSettings["password"];
            //string Domain = ConfigurationManager.AppSettings["domain"];
            //string WebUrl = "https://awsit.avivaworld.com";
            //string WebPartType = "CommentBoardWebPart";
            //String OutPutDirectory = @"F:\481921\ReplaceWebPart\out\list\replace\r";

            string SharePointOnline_OR_OnPremise = "OP";
            string UserName =ConfigurationManager.AppSettings["userName"];
            string Password = ConfigurationManager.AppSettings["password"];
            string Domain = ConfigurationManager.AppSettings["domain"];
            string WebUrl = ConfigurationManager.AppSettings["siteurl"];
            string WebPartType = "CommentBoardWebPart";
            String OutPutDirectory = ConfigurationManager.AppSettings["outputpath"];


            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper("GetWebPartUsage");
            if (SharePointOnline_OR_OnPremise.ToUpper().Equals("OP"))
            {
                webPartTransformationHelper.UseNetworkCredentialsAuthentication(UserName, Password, Domain);
            }
            else if (SharePointOnline_OR_OnPremise.ToUpper().Equals("OL"))
            {
                webPartTransformationHelper.UseOffice365Authentication(UserName, Password);
            }
            webPartTransformationHelper.AddSite(WebUrl);
            webPartTransformationHelper.WebPartType = WebPartType;
            webPartTransformationHelper.ExpandSubSites = true;
            webPartTransformationHelper.OutPutDirectory = OutPutDirectory;
            webPartTransformationHelper.Run();
        }

        public static void GetWebPartProperties()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string WebUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";
            string pageUrl = "/sites/TestContosoPublishingSite/Pages/TestWebPartTransformation.aspx";
            string outPutDir = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\GetWebPartProperties";
            //string webPartID = "d72fd359-2294-4434-b77c-514cc8ebafc5";
            string webPartID = "c7d8da93-0d50-4587-89d6-cd4e3663c660";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.GetWebPartProperties(pageUrl, webPartID,"N/A", WebUrl, outPutDir, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void GetWebPartProperties_UsingCSV()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string outPutDir = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\GetWebPartPropertiesByCSV\Current";
            string usageFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\GetWebPartPropertiesByCSV\WebParts_Usage.csv";
            string sourceWebPartTitle = "Featured Links Web Part";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.GetWebPartProperties_UsingCSV(sourceWebPartTitle, usageFilePath, outPutDir, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void ConfigureNewWebPartXml()
        {
            //string targetWebPartXmlFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\TargetContentQuery.xml";
            string targetWebPartXmlFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\AppPartXml.xml";
            string sourceXmlFilesDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\SourceDir";
            //string targetXmlFilesDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\TargetDir";
            string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.ConfigureNewWebPartXml(targetWebPartXmlFilePath, sourceXmlFilesDirectory, outPutDirectory);

        }

        public static void ConfigureNewWebPartXmlBulk()
        {
            string targetWebPartXmlFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Bulk\TargetContentQuery.xml";
            string sourceXmlFilesDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Bulk\SourceDir";
            //string targetXmlFilesDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Bulk\TargetDir";
            string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Bulk";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.ConfigureNewWebPartXml(targetWebPartXmlFilePath, sourceXmlFilesDirectory, outPutDirectory);

        }

        public static void DeleteWebPart()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            //string webUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";
            //string ServerRelativePageUrl = "/sites/TestContosoPublishingSite/Pages/TestWebPartTransformation.aspx";
            string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\DeleteWebPart";
            //string webPartTitle = "Featured Links Web Part";
            //string zoneIndex = "1";
            //string zoneId = "Left Column";
            //Guid webPartID = new Guid("c7d8da93-0d50-4587-89d6-cd4e3663c660");

            Guid StorageKey = new Guid("d2d4f7bd-6f2d-4677-b4f9-f63a59e09317");
            Guid webPartID = new Guid("40d61b73-b96e-44d4-a2d3-504dce482f99");
            string webUrl = "https://intranet.poc.com/sites/TestContosoTeamSite/";
            string ServerRelativePageUrl = "/sites/TestContosoTeamSite/SitePages/TestWebPart.aspx";
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.DeleteWebPart(webUrl, ServerRelativePageUrl, webPartID, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void DeleteWebPart_UsingCSV()
        {
            //string SharePointOnline_OR_OnPremise = "OP";
            //string UserName = "ms-mla-suamso";
            //string Password = "Password123";
            //string Domain = "MGMT7";
           // string usageFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\GetWebPartPropertiesByCSV\WebParts_Usage.csv";
            //string sourceWebPartTitle = "Featured Links Web Part";
           // string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\DeleteWebPartByCSV";



            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = ConfigurationManager.AppSettings["userName"];
            string Password = ConfigurationManager.AppSettings["password"];
            string Domain = ConfigurationManager.AppSettings["domain"];
            string WebUrl = ConfigurationManager.AppSettings["siteurl"];
            string outPutDirectory = ConfigurationManager.AppSettings["outputpath"];
            string usageFilePath = ConfigurationManager.AppSettings["usagefilepath"];
            string sourceWebPartTitle = ConfigurationManager.AppSettings["sourcewebparttype"];
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.DeleteWebPart_UsingCSV(sourceWebPartTitle, usageFilePath, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void AddWebPart_UsingCSV()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string usageFilePath = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\GetWebPartPropertiesByCSV\WebParts_Usage.csv";
            string sourceWebPartTitle = "Featured Links Web Part";
            //string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\AddWebPartByCsv";

            string configuredWebPartFileName = "ContentQuery.webpart";
            string configuredWebPartXmlDir = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\bulk\TargetDir";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.AddWebPart_UsingCSV(sourceWebPartTitle, configuredWebPartFileName, configuredWebPartXmlDir, usageFilePath, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void AddWebPart()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string webUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";
            string serverRelativePageUrl = "/sites/TestContosoPublishingSite/Pages/TestWebPartTransformation.aspx";
            string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\AddWebPart";
            string webPartZoneIndex = "0";
            string webPartZoneID = "LeftColumnZone";
            string configuredWebPartFileName = "N/A";
            string configuredWebPartXmlFile = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Single\TargetConfiguredWebPartXmls\Configured_c7d8da93-0d50-4587-89d6-cd4e3663c660_AppPartXml.xml";

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();

            webPartTransformationHelper.AddWebPart(webUrl, configuredWebPartFileName, configuredWebPartXmlFile, webPartZoneIndex, webPartZoneID, serverRelativePageUrl, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }

        public static void ReplaceWebPart()
        {
            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = "ms-mla-suamso";
            string Password = "Password123";
            string Domain = "MGMT7";
            string webUrl = "https://intranet.poc.com/sites/TestContosoPublishingSite/";
            string serverRelativePageUrl = "/sites/TestContosoPublishingSite/Pages/TestWebPartTransformation.aspx";
            string outPutDirectory = @"F:\481921\ReplaceWebPart\replace";
            string webPartZoneIndex = "0";
            string webPartZoneID = "TopLeftRow";
            string targetWebPartFileName = "Haveyoursay.dwp";
            string targetWebPartXmlFile = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ConfigureNewWebPartXml\Bulk\TargetContentQuery.xml";
            Guid sourceWebPartID = new Guid("b3d55ff9-d032-45e1-9670-cd442ba5cab3");

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.ReplaceWebPart(webUrl, targetWebPartFileName, targetWebPartXmlFile, sourceWebPartID, webPartZoneIndex, webPartZoneID, serverRelativePageUrl, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

        }

        public static void ReplaceWebPartByCSV()
        {
            //string SharePointOnline_OR_OnPremise = "OP";
            //string UserName = "sp_farm";
            //string Password = "SPfarm2013";
            //string Domain = "awdev";
            ////string outPutDirectory = @"E:\v-suamso\FTC-CAMSimulation\TransformationTool\TestingResults\ReplaceWebPartByCSV";
            //string targetWebPartFileName = "MSContentEditor.dwp";
            //string targetWebPartXmlDir = @"F:\481921\ReplaceWebPart\out\list\replace";
            //string sourceWebPartTitle = "CommentBoardWebPart";
            //string usageFilePath = @"F:\481921\ReplaceWebPart\out\list\replace\WebPartUsage.csv";

            string SharePointOnline_OR_OnPremise = "OP";
            string UserName = ConfigurationManager.AppSettings["userName"];
            string Password = ConfigurationManager.AppSettings["password"];
            string Domain = ConfigurationManager.AppSettings["domain"];
            string outPutDirectory = ConfigurationManager.AppSettings["outputpath"];
            string targetWebPartFileName = "SiteFeed.dwp";
            string targetWebPartXmlDir = ConfigurationManager.AppSettings["targetxmlpath"];
            string sourceWebPartTitle = "CommunityActivityFeedWebPart";
            string usageFilePath = ConfigurationManager.AppSettings["usagefilepath"];

            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.ReplaceWebPart_UsingCSV(sourceWebPartTitle, targetWebPartFileName, targetWebPartXmlDir, usageFilePath, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }
}
