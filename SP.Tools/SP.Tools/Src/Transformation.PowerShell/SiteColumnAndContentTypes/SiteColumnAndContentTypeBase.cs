using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    public class SiteColumnBase : Elementbase
    {
        public string Old_SiteColumn_Title { get; set; }
        public string Old_SiteColumn_InternalName { get; set; }
        public string Old_SiteColumn_ID { get; set; }
        public string Old_SiteColumn_Type { get; set; }
        public string old_SiteColumn_Scope { get; set; }

        public string New_SiteColumn_Title { get; set; }
        public string New_SiteColumn_InternalName { get; set; }
        public string New_SiteColumn_ID { get; set; }
        public string New_SiteColumn_Type { get; set; }
        public string New_SiteColumn_Scope { get; set; }
    }
    public class SiteColumnInput : Inputbase
    {
        public string CustomFields_Title { get; set; }
        public string CustomFields_InternalName { get; set; }
        public string CustomFields_Id { get; set; }
        public string CustomFields_Type { get; set; }
        public string CustomFields_Scope { get; set; }
        public string CustomFields_ListTitle { get; set; }
    }
    public class AddSiteColumnToContentTypeBase : Elementbase
    {
        public string ContentTypeName { get; set; }
        public string SiteColumnName { get; set; }

        public string ContentTypeID { get; set; }
        public string SiteColumnID { get; set; }
    }
       
    public class UpdateContentTypeinListBase : Elementbase
    {
        public string ListName { get; set; }
        public string oldContentTypeId { get; set; }
        public string newContentTypeName { get; set; }
    }
    public class UpdateContentTypeinListInput : Inputbase
    {
        public string ListName { get; set; }
        public string oldContentTypeId { get; set; }
        public string newContentTypeName { get; set; }
    }

    public class ContentTypeBase : Elementbase
    {
        public string OldContentTypeID { get; set; }
        public string OldContentTypeName { get; set; }
        public string NewContentTypeID { get; set; }
        public string NewContentTypeName { get; set; }
        
    }
    public class ContentTypeInput : Inputbase
    {
        public string ContentTypeId { get; set; }
        public string ContentTypeName { get; set; }
    }
}
