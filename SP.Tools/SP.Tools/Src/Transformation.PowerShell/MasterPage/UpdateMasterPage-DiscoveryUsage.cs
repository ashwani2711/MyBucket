using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.MasterPage
{
    [Cmdlet(VerbsCommon.Set, "MasterPageUsingDiscoveryUsage")]
    public class UpdateMasterPageUsingDiscoveryUsage : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string InputFolder;
        [Parameter(Mandatory = true, Position = 1)]
        public string New_MasterPageURL;
        [Parameter(Mandatory = true, Position = 2)]
        public string Old_MasterPageURL;
        [Parameter(Mandatory = true, Position = 3)]
        public string SharePointOnline_OR_OnPremise;
        [Parameter(Mandatory = true, Position = 4)]
        public string UserName;
        [Parameter(Mandatory = true, Position = 5)]
        public string Password;
        [Parameter(Mandatory = true, Position = 6)]
        public string Domain;
        
        protected override void ProcessRecord()
        {
            MasterPageHelper objMasterHelper = new MasterPageHelper();
            objMasterHelper.ChangeMasterPageForDiscoveryOutPut(InputFolder, New_MasterPageURL, Old_MasterPageURL, SharePointOnline_OR_OnPremise, UserName, Password, Domain);   
        }
    }
}
