﻿using PnP.PowerShell.Commands.Attributes;
using PnP.PowerShell.Commands.Base;
using PnP.PowerShell.Commands.Base.PipeBinds;
using PnP.PowerShell.Commands.Utilities;
using System.Management.Automation;
using Group = PnP.PowerShell.Commands.Model.Graph.Group;

namespace PnP.PowerShell.Commands.EntraID
{
    [Cmdlet(VerbsCommon.Remove, "PnPEntraIDGroup")]
    [RequiredMinimalApiPermissions("Group.ReadWrite.All")]
    [Alias("Remove-PnPAzureADGroup")]
    public class RemoveEntraIDGroup : PnPGraphCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true)]
        public EntraIDGroupPipeBind Identity;

        protected override void ExecuteCmdlet()
        {
            if (Identity != null)
            {
                Group group = Identity.GetGroup(Connection, AccessToken);

                if (group != null)
                {
                    Microsoft365GroupsUtility.RemoveGroupAsync(Connection, new System.Guid(group.Id), AccessToken).GetAwaiter().GetResult();
                }
            }
        }
    }
}