using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using System.Data.SqlClient;
using Microsoft.Xrm.Sdk;
using GsmComm.GsmCommunication;
using Microsoft.Xrm.Sdk.Workflow;
using GsmComm.PduConverter;
using GsmComm.PduConverter.SmartMessaging;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace Sigma_SendOutgoingSms
{
    ///  <summary>	
	///  - This project is for send sms via GSM Modem
	///  </summary>    

    public sealed class SendOutgoingSms : CodeActivity
    {
        ///  <summary>	
        ///  SmsMustSend is SMS Message created in CRM and must be send by GSM Modem
        ///  SmsQueue is Queue for SMS that this SMS is in it
        ///  </summary> 
        [Input("SMS")]
        [ReferenceTarget("new_sms")]
        public InArgument<EntityReference> SmsMustSend { get; set; }

        [Input("Sms Queue")]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> SmsQueue { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the workflow context and CRM Service
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService crmService = serviceFactory.CreateOrganizationService(context.UserId);
        
            //Get ID for SMS Message
            Guid smsId = SmsMustSend.Get(executionContext).Id;

            //Get ID for SMS Queue
            Guid smsQueueId = SmsQueue.Get(executionContext).Id;

            //Get SMS Message info (to,from,cc,message) by SMS ID
            Entity sms = crmService.Retrieve("new_sms", smsId, new ColumnSet("to", "from","cc", "new_message"));

            //Get SMS to info
            Entity to = ((EntityCollection)sms.Attributes["to"]).Entities[0];

            //Get SMS cc info
            List<Entity> cc = sms.Contains("cc") ? ((EntityCollection)sms.Attributes["cc"]).Entities.ToList<Entity>() : null;

            //Get ID for Activity Party for SMS to recipient
            Guid toId = ((EntityReference)to.Attributes["partyid"]).Id;

            //Get recipient Contact info by ID
            Entity contact = crmService.Retrieve("contact", toId, new ColumnSet(true));

            //Get recipient mobile number
            string mobile = contact.Contains("mobilephone") ? contact.Attributes["mobilephone"].ToString() : "";

            //Get SMS message body
            string message = sms.Attributes["new_message"].ToString();

            if(mobile != "")
            {
                //Call Send SMS Method for sending SMS with GSM Modem
                SendSms(message, mobile);

                #region Send SMS To CC
                if (cc.Count() > 0)
                {
                    foreach(var recipient in cc)
                    {
                        //Get ID for Activity Party for SMS cc recipient
                        Guid recipientId = ((EntityReference)recipient.Attributes["partyid"]).Id;

                        //Get Entity Logical Name for Activity Party for SMS cc recipient
                        string recipientType = ((EntityReference)recipient.Attributes["partyid"]).LogicalName;

                        //Get cc recipient mobile number
                        string recipientMobile = SigmaXrm.SMS.GetMobileById(recipientId, recipientType, crmService);

                        if (recipientMobile != "")
                        {
                            //Call SendSms Method for Send SMS to cc with GSM Modem
                            SendSms(message, recipientMobile);
                        }
                            
                    }
                }
                #endregion

                //Add SMS Message to SMS Queue for Process
                SigmaXrm.Queue.AddItemToQueue(smsId, "new_sms", smsQueueId, crmService);

                //Update SMS Delivery Status
                SigmaXrm.SMS.UpdateSmsDeliveryOn(smsId, crmService);
            }
        }

        ///  <summary>	
        ///  This Method for Send SMS Message with GSM Modem
        ///  </summary>    
        public void SendSms(string message,string mobile)
        {
            //Get GSM Modem Config
            List<SigmaXrm.Credentials> credentials = SigmaXrm.Report.GetServerConfig();

            //Set GSM Modem Port Number
            string portNumber = credentials.First(p => p.Label == "SMS_Port_Number").Value.ToString();

            //Set GSM Modem Baudrate
            string baudRate = credentials.First(p => p.Label == "SMS_Port_Baud_Rate").Value.ToString();

            //Set GSM Modem Timeout
            string timeout = credentials.First(p => p.Label == "SMS_Port_Timeout").Value.ToString();

            //Create GSM Modem Object
            GsmCommMain comm = new GsmCommMain(portNumber, Convert.ToInt32(baudRate), Convert.ToInt32(timeout));

            if (!comm.IsOpen())
            {
                try
                {
                    //Open GSM Modem Port
                    comm.Open();
                }
                catch (Exception f)
                {
                    if (comm.IsOpen())
                        comm.Close();
                }
            }

            if (comm.IsOpen())
            {
                #region Send SMS
                try
                {
                    //Send SMS Message with GSM Modem
                    SmsSubmitPdu[] pdu_list = SmartMessageFactory.CreateConcatTextMessage(message, true, mobile);
                    comm.SendMessages(pdu_list);
                    comm.Close();
                }
                catch (Exception ex)
                {                   
                    if (comm.IsOpen())
                        comm.Close();
                }
                #endregion
            }
        }
    }
}
