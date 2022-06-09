using System.Management.Automation;
using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Base.PipeBinds;

namespace PnP.PowerShell.Commands.Apps
{
    [Cmdlet(VerbsCommon.Get, "PnPAppInstance")]
    public class GetAppInstance : PnPWebRetrievalsCmdlet<AppInstance>
    {
        [Parameter(Mandatory = false, Position=0, ValueFromPipeline = true, HelpMessage = "Specifies the Id of the App Instance")]
        public AppPipeBind Identity;

        protected override void ExecuteCmdlet()
        {
            if (Identity != null)
            {
                var instance = Identity.GetAppInstance(CurrentWeb);
                WriteObject(instance);
            }
            else
            {
                var instances = CurrentWeb.GetAppInstances();
                if (instances.Count > 1)
                {
                    WriteObject(instances, true);
                }
                else if (instances.Count == 1)
                {
                    WriteObject(instances[0]);
                }
            }
        }
    }
}