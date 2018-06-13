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
    public class FdxRollupCrOnLeadStateChange : IPlugin
    {
        int step = 0;
        IOrganizationService service;
        ITracingService tracingService;

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                step = 1;
                tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

                step = 2;
                service = serviceFactory.CreateOrganizationService(context.UserId);

                step = 3;
                WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());


                ParameterCollection contextInputParameter = context.InputParameters;
                EntityReference leadEntityReference = new EntityReference();
                EntityReference parentContactEntityReference = new EntityReference();
                EntityReference parentAccountEntityReference = new EntityReference();
                EntityReference parentOpportunityEntityReference = new EntityReference();
                
                Entity leadEntiy = new Entity();


                tracingService.Trace("Connected Lead -> ROll-up CR on Lead stage change");

                step = 4;
                if (contextInputParameter.Contains("LeadId")) //On Lead Qualify....
                {
                    if (contextInputParameter["LeadId"] is EntityReference)
                    {
                        tracingService.Trace("Connected Lead -> ROll-up CR on Lead Qualify");
                        step = 100;

                        leadEntityReference = (EntityReference)context.InputParameters["LeadId"];

                        step = 101;
                        tracingService.Trace("conetxt dept - " + context.Depth.ToString());
                        tracingService.Trace("logical name - " + leadEntityReference.LogicalName);
                        
                        if (leadEntityReference.LogicalName != "lead" || context.Depth != 1)
                            return;

                        step = 102;
                        step = 103;
                        leadEntiy = this.getLead(leadEntityReference);

                        #region Set parent account and contact reference against the lead....
                        if (leadEntiy.Attributes.Contains("parentcontactid"))
                            parentContactEntityReference = (EntityReference)leadEntiy.Attributes["parentcontactid"];

                        if (leadEntiy.Attributes.Contains("parentaccountid"))
                            parentAccountEntityReference = (EntityReference)leadEntiy.Attributes["parentaccountid"];
                        #endregion

                        step = 104;
                        #region Get reference to the entities that got created durinng lead qualification....
                        if (context.OutputParameters.Contains("CreatedEntities"))
                        {
                            foreach (EntityReference crEntities in ((IEnumerable)context.OutputParameters["CreatedEntities"]))
                            {
                                step = 105;
                                switch (crEntities.LogicalName)
                                {
                                    case "opportunity":
                                        parentOpportunityEntityReference = crEntities;
                                        break;
                                    case "account":
                                        parentAccountEntityReference = crEntities;
                                        break;
                                    case "contact":
                                        parentContactEntityReference = crEntities;
                                        break;
                                }
                            }
                        }
                        #endregion

                        step = 106;
                        this.rollupCampaignResponse(leadEntityReference, parentContactEntityReference, parentAccountEntityReference, parentOpportunityEntityReference);
                    }
                }
                else if (contextInputParameter.Contains("EntityMoniker")) //On Lead DisQualify...
                {
                    if (contextInputParameter["EntityMoniker"] is EntityReference)
                    {
                        tracingService.Trace("Connected Lead -> ROll-up CR on Lead DisQualify");
                        step = 200;

                        var leadState = (OptionSetValue)contextInputParameter["State"];

                        leadEntityReference = (EntityReference)contextInputParameter["EntityMoniker"];

                        step = 201;
                        tracingService.Trace("conetxt dept - " + context.Depth.ToString());
                        tracingService.Trace("logical name - " + leadEntityReference.LogicalName);
                        tracingService.Trace(string.Format("{0}", (leadEntityReference.LogicalName != "lead" || context.Depth > 2)));
                        if (leadEntityReference.LogicalName != "lead" || context.Depth > 2)
                            return;

                        step = 202;
                        if (leadState.Value == 2)
                        {
                            step = 203;
                            leadEntiy = this.getLead(leadEntityReference);

                            #region Set parent account and contact reference against the lead....
                            if (leadEntiy.Attributes.Contains("parentcontactid"))
                                parentContactEntityReference = (EntityReference)leadEntiy.Attributes["parentcontactid"];

                            if (leadEntiy.Attributes.Contains("parentaccountid"))
                                parentAccountEntityReference = (EntityReference)leadEntiy.Attributes["parentaccountid"];
                            #endregion

                            step = 204;
                            this.rollupCampaignResponse(leadEntityReference, parentContactEntityReference, parentAccountEntityReference);
                        }
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(string.Format("An error occurred in the FdxRollupCrOnLeadStateChange plug-in at Step {0}.", step), ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("FdxRollupCrOnLeadStateChange: step {0}, {1}", step, ex.ToString());
                throw;
            }
        }

        private Entity getLead(EntityReference _leadEntityReference)
        {
            step = 300;
            Entity lead = service.Retrieve("lead", _leadEntityReference.Id, new ColumnSet("leadid", "parentaccountid", "parentcontactid"));

            return lead;
        }

        private void rollupCampaignResponse(EntityReference _leadEntityReference, EntityReference _parentContactEntityReference, EntityReference _parentAccountEntityReference, EntityReference _parentOpportunityEntityReference = null)
        {
            try
            {
                step = 400;
                EntityCollection campaignResponseCollection = this.getCampaignResponses(_leadEntityReference);

                foreach (Entity campaignResponse in campaignResponseCollection.Entities)
                {
                    step = 401;

                    if (!campaignResponse.Attributes.Contains("fdx_reconversioncontact") && _parentContactEntityReference != null)
                    {
                        campaignResponse.Attributes["fdx_reconversioncontact"] = new EntityReference("contact", _parentContactEntityReference.Id);
                    }

                    step = 402;
                    if (!campaignResponse.Attributes.Contains("fdx_reconversionopportunity") && _parentOpportunityEntityReference != null)
                    {
                        campaignResponse.Attributes["fdx_reconversionopportunity"] = new EntityReference("opportunity", _parentOpportunityEntityReference.Id);
                    }

                    step = 404;
                    if (!campaignResponse.Attributes.Contains("customer") && _parentAccountEntityReference != null)
                    {
                        Entity customer = new Entity("activityparty");
                        customer.Attributes["partyid"] = _parentAccountEntityReference;
                        EntityCollection Customerentity = new EntityCollection();
                        Customerentity.Entities.Add(customer);

                        campaignResponse["customer"] = Customerentity;
                    }

                    //Update crm rollup boolean...
                    campaignResponse.Attributes["fdx_crrollup"] = true;

                    step = 405;
                    service.Update(campaignResponse);
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(string.Format("An error occurred in the FdxRollupCrOnLeadStateChange plug-in at Step {0}.", step), ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("FdxRollupCrOnLeadStateChange: step {0}, {1}", step, ex.ToString());
                throw;
            }
        }

        private EntityCollection getCampaignResponses(EntityReference _leadEntityReference)
        {
            step = 500;
            EntityCollection campaignResponseCollection = new EntityCollection();

            step = 501;
            QueryExpression queryExp = CRMQueryExpression.getQueryExpression("campaignresponse", new ColumnSet("fdx_reconversionlead", "fdx_reconversioncontact", "fdx_reconversionopportunity", "customer"), new CRMQueryExpression[] { new CRMQueryExpression("fdx_reconversionlead", ConditionOperator.Equal, _leadEntityReference.Id) }, LogicalOperator.And);

            step = 502;
            campaignResponseCollection = service.RetrieveMultiple(queryExp);

            return campaignResponseCollection;
        }
    }
}
