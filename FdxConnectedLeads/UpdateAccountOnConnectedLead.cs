using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxConnectedLeads
{
    /// <summary>
    /// SMART-821: Tag account to connected lead with context lead's account if empty
    /// </summary>
    public class UpdateAccountOnConnectedLead
    {
        public static void updateConnectLeadAcc(Entity _entity, Guid _contextLead, IOrganizationService _service)
        {
            Entity lead = _service.Retrieve("lead", _entity.Id, new ColumnSet(true));

            if (!lead.Attributes.Contains("parentaccountid"))
            {
                lead.Attributes["parentaccountid"] = new EntityReference("account", _contextLead);
            }

            _service.Update(lead);
        }
    }
}
