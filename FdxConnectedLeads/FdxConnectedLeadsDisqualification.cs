using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxConnectedLeads
{
    public class FdxConnectedLeadsDisqualification : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution context from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            int step = 0;

            //Call Input parameter collection to get all the data passes....

            //on Update Leaddisqualify message 

            //On Lead Qualify message
            if (context.InputParameters.Contains("EntityMoniker") && context.InputParameters["EntityMoniker"] is EntityReference)
            {

                EntityReference leadEntityReference = (EntityReference)context.InputParameters["EntityMoniker"];
                var state = (OptionSetValue)context.InputParameters["State"];
                var status = (OptionSetValue)context.InputParameters["Status"];


                if (leadEntityReference.LogicalName != "lead" || context.Depth != 1)
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    step = 0;
                    if (leadEntityReference.LogicalName == "lead" && state.Value == 2)
                    {
                        step = 12;
                        #region Fetch groupid of context lead
                        QueryExpression contextLeadQuery = new QueryExpression();
                        contextLeadQuery.EntityName = "lead";
                        contextLeadQuery.ColumnSet = new ColumnSet("fdx_groupid", "leadid");
                        contextLeadQuery.Criteria.AddCondition("leadid", ConditionOperator.Equal, leadEntityReference.Id);
                        EntityCollection contextLeadCollection = service.RetrieveMultiple(contextLeadQuery);
                        #endregion

                        step = 1;
                        if (contextLeadCollection.Entities.Count > 0)
                        {
                            Entity contextLead = contextLeadCollection[0];
                            step = 2;
                            #region Fetch Connected Leads except context lead(Leads with similar Group Id)
                            QueryExpression connectedLeadsQuery = new QueryExpression();
                            connectedLeadsQuery.EntityName = "lead";
                            connectedLeadsQuery.ColumnSet = new ColumnSet("leadid");
                            connectedLeadsQuery.Criteria.AddFilter(LogicalOperator.And);
                            connectedLeadsQuery.Criteria.AddCondition("fdx_groupid", ConditionOperator.Equal, contextLead.Attributes["fdx_groupid"]);
                            connectedLeadsQuery.Criteria.AddCondition("leadid", ConditionOperator.NotEqual, contextLead.Id);
                            connectedLeadsQuery.Criteria.AddCondition("statecode", ConditionOperator.NotEqual, 2);
                            //and lead is not closed0
                            EntityCollection connectedLeadsCollection = service.RetrieveMultiple(connectedLeadsQuery);
                            #endregion
                            step = 3;
                            foreach (Entity connectedLead in connectedLeadsCollection.Entities)
                            {
                                #region Disqualify connected lead
                                step = 4;
                                SetStateRequest request = new SetStateRequest
                                {
                                    EntityMoniker = new EntityReference("lead", connectedLead.Id),
                                    State = new OptionSetValue(2),    //Status = disqualify(2)
                                    Status = new OptionSetValue(756480016) //Status Reason = disqualify - connected lead  
                                };
                                step = 15;
                                service.Execute(request);
                                step = 13;
                                #endregion
                            }
                        }
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("Plugin Error: " + ex.Message + ". Exception occurred at step = {0}.", step));
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("Plugin Error: " + ex.Message + ". Exception occurred at step = {0}.", step));
                }
            }
        }
    }
}
