using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;

namespace Sigma_SetPollEmailRegarding
{
    ///  <summary>	
	///  - This project is for connect poll email to related incident
	///  </summary>    
    public sealed class SetPollEmailRegarding : CodeActivity
    {
        ///  <summary>	
        ///  PollEmail is Poll email recieved and we want to connect it to related incident
        ///  </summary> 
        [Input("Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> PollEmail { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService crmService = serviceFactory.CreateOrganizationService(context.UserId);

            //Get ID for Poll Email
            Guid emailId = PollEmail.Get(executionContext).Id;

            try
            {
                //get current email
                Entity current_email_entity = crmService.Retrieve("email", emailId, new ColumnSet(true));

                //get email subject
                string email_subject = current_email_entity.Attributes["subject"].ToString();

                //extract case number from poll email subject
                string related_case_number_by_email_subject = email_subject.Substring(email_subject.IndexOf("CAS")).Trim();

                //search for related incident by case number
                List<Entity> relatedIncident = SigmaXrm.Search.SearchInCrmWithSingleCondition("incident", "ticketnumber", ConditionOperator.Equal, related_case_number_by_email_subject, crmService);

                if (relatedIncident.Count() > 0)
                {
                    //Get ID for related incident
                    Guid incidentId = new Guid(relatedIncident[0].Attributes["incidentid"].ToString());

                    //Set Regarding For Poll Email and connect email to founded incident
                    Entity poll_email_for_update_regarding = new Entity("email");
                    poll_email_for_update_regarding.Id = emailId;
                    poll_email_for_update_regarding.Attributes["regardingobjectid"] = new EntityReference("incident", incidentId);
                    crmService.Update(poll_email_for_update_regarding);
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in the Set Regarding For Poll Mail plug-in. ", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("Set Regarding For Poll Mail: {0}", ex.ToString());
                throw;
            }
        }
    }
}
