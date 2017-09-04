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
    public class FdxConnectedLeadsQualification : IPlugin
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
            if (context.InputParameters.Contains("LeadId") && context.InputParameters["LeadId"] is EntityReference)
            {
                EntityReference leadEntityReference = (EntityReference)context.InputParameters["LeadId"];

                if (leadEntityReference.LogicalName != "lead" || context.Depth != 1)
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    step = 0;

                    #region Fetch groupid of context lead
                    QueryExpression contextLeadQuery = new QueryExpression();
                    contextLeadQuery.EntityName = "lead";
                    contextLeadQuery.ColumnSet = new ColumnSet("fdx_groupid", "leadid", "statuscode", "statecode", "parentaccountid", "parentcontactid","fdx_isdecisionmaker");
                    contextLeadQuery.Criteria.AddCondition("leadid", ConditionOperator.Equal, leadEntityReference.Id);
                    EntityCollection contextLeadCollection = service.RetrieveMultiple(contextLeadQuery);
                    #endregion

                    step = 115;
                    #region Retrive stakeholder & decisionmaker connectionroleid to update opportunity stakeholders connection role
                    QueryExpression connectionRoleQuery = new QueryExpression();
                    connectionRoleQuery.EntityName = "connectionrole";
                    connectionRoleQuery.ColumnSet = new ColumnSet("connectionroleid", "name");
                    connectionRoleQuery.Criteria.AddCondition("name", ConditionOperator.In, new string[] { "Stakeholder", "Decision Maker" });
                    EntityCollection connectionRoleCollection = service.RetrieveMultiple(connectionRoleQuery);

                    Guid conrole_StakeholderId = Guid.Empty;
                    Guid conrole_DecisionMakerId = Guid.Empty;
                    foreach (Entity connectionRole in connectionRoleCollection.Entities)
                    {
                        step = 4;
                        switch (connectionRole.Attributes["name"].ToString())
                        {
                            case "Stakeholder":
                                conrole_StakeholderId = connectionRole.Id;
                                break;
                            case "Decision Maker":
                                //newAccount = service.Retrieve(crEntities.LogicalName, crEntities.Id, new ColumnSet("accountid"));
                                conrole_DecisionMakerId = connectionRole.Id;
                                break;

                        }
                    }
                    #endregion

                    step = 1;
                    if (contextLeadCollection.Entities.Count > 0)
                    {
                        Entity contextLead = contextLeadCollection[0];

                        #region Fetch Connected Leads except context lead(Leads with similar Group Id)
                        QueryExpression connectedLeadsQuery = new QueryExpression();
                        connectedLeadsQuery.EntityName = "lead";
                        connectedLeadsQuery.ColumnSet = new ColumnSet("fdx_groupid", "leadid", "parentcontactid", "firstname", "lastname", "telephone1", "telephone2", "fdx_salutation", "fdx_credential", "fdx_jobtitlerole", "emailaddress1", "address1_city", "fdx_stateprovince", "fdx_zippostalcode", "address1_country", "address1_line1", "address1_line2","fdx_isdecisionmaker");//add all contact attributes***
                        connectedLeadsQuery.Criteria.AddFilter(LogicalOperator.And);
                        connectedLeadsQuery.Criteria.AddCondition("fdx_groupid", ConditionOperator.Equal, contextLead.Attributes["fdx_groupid"]);
                        connectedLeadsQuery.Criteria.AddCondition("leadid", ConditionOperator.NotEqual, contextLead.Id);
                        connectedLeadsQuery.Criteria.AddCondition("statecode", ConditionOperator.NotEqual, 2);
                        //and lead is not closed0
                        EntityCollection connectedLeadsCollection = service.RetrieveMultiple(connectedLeadsQuery);
                        #endregion

                        #region Retrive Account, Contact & Opportunity created on Qualify
                        step = 2;
                        //Entity opportunity = null;
                        //Entity newAccount = null;
                        //Entity newContact = null;
                        Guid qual_opportunityId = Guid.Empty;
                        Guid qual_accountId = Guid.Empty;
                        Guid qual_contactId = Guid.Empty;
                        if (context.OutputParameters.Contains("CreatedEntities"))
                        {
                            step = 3;
                            foreach (EntityReference crEntities in ((IEnumerable)context.OutputParameters["CreatedEntities"]))
                            {
                                step = 4;
                                switch (crEntities.LogicalName)
                                {
                                    case "opportunity":
                                        qual_opportunityId = crEntities.Id;
                                        //opportunity = service.Retrieve(crEntities.LogicalName, crEntities.Id, new ColumnSet("opportunityid"));
                                        break;
                                    case "account":
                                        //newAccount = service.Retrieve(crEntities.LogicalName, crEntities.Id, new ColumnSet("accountid"));
                                        qual_accountId = crEntities.Id;
                                        break;
                                    case "contact":
                                        //newContact = service.Retrieve(crEntities.LogicalName, crEntities.Id, new ColumnSet("contactid"));
                                        qual_contactId = crEntities.Id;
                                        break;
                                }
                            }
                        }
                        #endregion

                        //Loop through connected leads
                        step = 5;
                        Entity contact;
                        foreach (Entity connectedLead in connectedLeadsCollection.Entities)
                        {
                            contact = null;
                            if (connectedLead.LogicalName == "lead")
                            {
                                step = 6;
                                //If connected lead existing contact doesnot exist
                                if (!connectedLead.Attributes.Contains("parentcontactid"))
                                {
                                    #region Create contact for connected lead
                                    //add all mapping attributes lead to contact***
                                    contact = new Entity("contact", Guid.NewGuid());
                                    step = 7;
                                    if (connectedLead.Attributes.Contains("firstname"))
                                    {
                                        contact.Attributes["firstname"] = connectedLead.Attributes["firstname"];
                                    }
                                    if (connectedLead.Attributes.Contains("lastname"))
                                    {
                                        contact.Attributes["lastname"] = connectedLead.Attributes["lastname"];
                                    }
                                    if (connectedLead.Attributes.Contains("telephone1"))
                                    {
                                        contact.Attributes["telephone1"] = connectedLead.Attributes["telephone1"];
                                    }
                                    if (connectedLead.Attributes.Contains("telephone2"))
                                    {
                                        contact.Attributes["telephone2"] = connectedLead.Attributes["telephone2"];
                                    }
                                    if (connectedLead.Attributes.Contains("fdx_salutation"))
                                    {
                                        contact.Attributes["fdx_salutation"] = connectedLead.Attributes["fdx_salutation"];
                                    }
                                    if (connectedLead.Attributes.Contains("fdx_credential"))
                                    {
                                        contact.Attributes["fdx_credential"] = connectedLead.Attributes["fdx_credential"];
                                    }
                                    if (connectedLead.Attributes.Contains("fdx_jobtitlerole"))
                                    {
                                        contact.Attributes["fdx_jobtitlerole"] = connectedLead.Attributes["fdx_jobtitlerole"];
                                    }
                                    if (connectedLead.Attributes.Contains("emailaddress1"))
                                    {
                                        contact.Attributes["emailaddress1"] = connectedLead.Attributes["emailaddress1"];
                                    }
                                    if (connectedLead.Attributes.Contains("address1_line1"))
                                    {
                                        contact.Attributes["address1_line1"] = connectedLead.Attributes["address1_line1"];
                                    }
                                    if (connectedLead.Attributes.Contains("address1_line2"))
                                    {
                                        contact.Attributes["address1_line2"] = connectedLead.Attributes["address1_line2"];
                                    }
                                    if (connectedLead.Attributes.Contains("address1_city"))
                                    {
                                        contact.Attributes["address1_city"] = connectedLead.Attributes["address1_city"];
                                    }
                                    if (connectedLead.Attributes.Contains("fdx_stateprovince"))
                                    {
                                        contact.Attributes["fdx_stateprovinceid"] = new EntityReference("fdx_state", ((EntityReference)connectedLead.Attributes["fdx_stateprovince"]).Id);
                                    }
                                    if (connectedLead.Attributes.Contains("fdx_zippostalcode"))
                                    {
                                        contact.Attributes["fdx_zippostalcodeid"] = new EntityReference("fdx_zipcode", ((EntityReference)connectedLead.Attributes["fdx_zippostalcode"]).Id);
                                    }
                                    if (connectedLead.Attributes.Contains("address1_country"))
                                    {
                                        contact.Attributes["address1_country"] = connectedLead.Attributes["address1_country"];
                                    }
                                    contact.Attributes["originatingleadid"] = new EntityReference("lead", connectedLead.Id);

                                    step = 8;
                                    //tag contact to account 
                                    if (contextLead.Attributes.Contains("parentaccountid"))
                                        contact.Attributes["parentcustomerid"] = new EntityReference("account", ((EntityReference)contextLead.Attributes["parentaccountid"]).Id);
                                    else if (qual_accountId != Guid.Empty)
                                        contact.Attributes["parentcustomerid"] = new EntityReference("account", qual_accountId);

                                    step = 9;
                                    service.Create(contact);
                                    #endregion

                                    #region Tag new contact to connected lead
                                    step = 10;
                                    connectedLead.Attributes["parentcontactid"] = new EntityReference("contact", contact.Id);
                                    service.Update(connectedLead);
                                    #endregion

                                }

                                #region Disqualify connected lead

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

                                #region Create a stakeholder connection in opportunity for connected lead -> new contact or existing contact
                                step = 11;
                                Entity connection = new Entity("connection");

                                if (contact != null)
                                {
                                    step = 113;
                                    connection.Attributes["record2id"] = new EntityReference("contact", contact.Id);
                                }
                                else if (connectedLead.Attributes.Contains("parentcontactid"))
                                {
                                    step = 114;
                                    connection.Attributes["record2id"] = new EntityReference("contact", ((EntityReference)connectedLead.Attributes["parentcontactid"]).Id);
                                }
                                connection.Attributes["record1id"] = new EntityReference("opportunity", qual_opportunityId);
                                
                                step = 112;
                                //if connected lead is decisionmaker, mark stakeholder as decisionmaker
                                if(connectedLead.Attributes.Contains("fdx_isdecisionmaker") && (bool)connectedLead.Attributes["fdx_isdecisionmaker"]== true)
                                    connection.Attributes["record2roleid"] = new EntityReference("connectionrole", conrole_DecisionMakerId);
                                else
                                    connection.Attributes["record2roleid"] = new EntityReference("connectionrole", conrole_StakeholderId);

                                step = 12;
                                service.Create(connection);
                                step = 16;
                                #endregion
                            }
                        }
                        #region (Commented)Change primary contact for context lead -> existing account
                        //if (contextLead.Attributes.Contains("parentaccountid"))
                        //{
                        //    Entity existingAccount = service.Retrieve("account", ((EntityReference)contextLead.Attributes["parentaccountid"]).Id, new ColumnSet("accountid", "primarycontactid"));

                        //    existingAccount.Attributes["primarycontactid"] = contextLead.Attributes.Contains("parentcontactid") ? new EntityReference("contact", ((EntityReference)contextLead.Attributes["parentcontactid"]).Id) : new EntityReference("contact", qual_contactId);

                        //    service.Update(existingAccount);
                        //}
                        //step = 14;
                        #endregion


                        #region Update Qualified lead Stakeholder connection role on opportunity
                        step = 31;
                        //retrive connection(stakeholder) created on qualify of Lead (for new contact/existing contact) under opportunity
                        QueryExpression qual_connectionQuery = new QueryExpression("connection");
                        qual_connectionQuery.ColumnSet = new ColumnSet("connectionid", "record2roleid");
                        qual_connectionQuery.Criteria.AddFilter(LogicalOperator.And);
                        qual_connectionQuery.Criteria.AddCondition("record1id", ConditionOperator.Equal, qual_opportunityId);
                        qual_connectionQuery.Criteria.AddCondition("record2id", ConditionOperator.Equal, contextLead.Attributes.Contains("parentcontactid") ? ((EntityReference)contextLead.Attributes["parentcontactid"]).Id : qual_contactId);

                        step = 32;
                        EntityCollection qual_connectionCollection = service.RetrieveMultiple(qual_connectionQuery);
                        Entity qual_connection = qual_connectionCollection[0];

                        step = 33;
                        //if context lead is decisionmaker, mark stakeholder as decisionmaker
                        if (contextLead.Attributes.Contains("fdx_isdecisionmaker") && (bool)contextLead.Attributes["fdx_isdecisionmaker"] == true)
                            qual_connection.Attributes["record2roleid"] = new EntityReference("connectionrole", conrole_DecisionMakerId);
                        else
                            qual_connection.Attributes["record2roleid"] = new EntityReference("connectionrole", conrole_StakeholderId);

                        service.Update(qual_connection);
                        step = 33;
                        #endregion
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
