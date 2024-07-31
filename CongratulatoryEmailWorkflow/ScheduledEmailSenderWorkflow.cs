using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;
using System.ServiceModel;

namespace CongratulatoryEmailWorkflow
{
    public class ScheduledEmailSenderWorkflow : CodeActivity
    {
        [Input("Receiver Guid")]
        [RequiredArgument]
        public InArgument<string> ReceiverGuid { get; set; }

        [Input("Sender Guid")]
        [RequiredArgument]
        public InArgument<string> SenderGuid { get; set; }

        [Input("Email Template Name")]
        [RequiredArgument]
        public InArgument<string> EmailTemplateName { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            IOrganizationServiceFactory factory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService organizationService = factory.CreateOrganizationService(context.InitiatingUserId);

            try
            {
                tracingService.Trace("Started scheduled email send activity.");

                // Initialize variables.
                string receiverGuid = null;
                string senderGuid = null;
                string emailTemplateName = null;
                Entity receiver = null;
                string receiverFirstname = null;
                string receiverLastname = null;
                int receiverGenderCode = 0;
                DateTime receiverBirthdate = DateTime.MinValue;

                // Check for missing fields.
                if (string.IsNullOrEmpty(ReceiverGuid.Get(executionContext)))
                {
                    tracingService.Trace("Receiver GUID is missing.");
                    throw new InvalidPluginExecutionException("Receiver GUID is missing.");
                }
                else
                    receiverGuid = ReceiverGuid.Get(executionContext);

                if (string.IsNullOrEmpty(SenderGuid.Get(executionContext)))
                {
                    tracingService.Trace("Sender GUID is missing.");
                    throw new InvalidPluginExecutionException("Sender GUID is missing.");
                }
                else
                    senderGuid = SenderGuid.Get(executionContext);

                if (string.IsNullOrEmpty(EmailTemplateName.Get(executionContext)))
                {
                    tracingService.Trace("Email template name is missing.");
                    throw new InvalidPluginExecutionException("Email template name is missing.");
                }
                else
                    emailTemplateName = EmailTemplateName.Get(executionContext);

                // Retrieve receiver data.
                receiver = organizationService.Retrieve("contact", new Guid(receiverGuid), new ColumnSet("firstname", "lastname", "birthdate", "gendercode"));

                // Check for missing receiver data.
                if (receiver == null)
                {
                    tracingService.Trace($"Contact with GUID {receiverGuid} not found.");
                    throw new InvalidPluginExecutionException($"Contact with GUID {receiverGuid} not found.");
                }

                receiverFirstname = receiver.GetAttributeValue<string>("firstname");
                receiverLastname = receiver.GetAttributeValue<string>("lastname");
                receiverBirthdate = receiver.GetAttributeValue<DateTime>("birthdate");
                OptionSetValue genderCodeOptionSet = receiver.GetAttributeValue<OptionSetValue>("gendercode");

                if (genderCodeOptionSet != null)
                     receiverGenderCode = genderCodeOptionSet.Value;

                tracingService.Trace($"GenderCode = {receiverGenderCode}.");

                // Send congratulatory email
                SendCongratulatoryEmail(emailTemplateName, receiverGuid, receiverFirstname, receiverLastname, receiverBirthdate, receiverGenderCode, senderGuid, organizationService, tracingService);

                tracingService.Trace("Ended scheduled email send activity.");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Congratulator Plugin: {0}.", ex.Message);  // Log exception.
                throw;
            }
        }

        private SendEmailResponse SendCongratulatoryEmail(string emailTemplateName, string receiverGuid, string receiverFirstname, string receiverLastname, DateTime receiverBirthdate, int genderCode, string senderGuid, IOrganizationService organizationService, ITracingService tracingService)
        {
            tracingService.Trace("Getting email template.");
            string emailTemplate = EmailConstants.EmailTemplates[emailTemplateName];
            tracingService.Trace("Finished getting email template.");

            string gendercodeTitle = string.Empty;

            switch (genderCode)
            {
                case 1: // Male.
                    gendercodeTitle = "Sehr geehrter Herr";
                    break;
                case 2: // Female.
                    gendercodeTitle = "Sehr geehrte Frau";
                    break;
            }

            // Replace the placeholders in the email body
            string body = emailTemplate.Replace("[Firstname]", receiverFirstname)
                                                   .Replace("[Lastname]", receiverLastname)
                                                   .Replace("[GendercodeTitle]", gendercodeTitle)
                                                   .Replace("[Birthdate]", receiverBirthdate.ToString("d"));

            Entity email = new Entity("email");

            Entity fromParty = new Entity("activityparty");
            Entity toParty = new Entity("activityparty");

            fromParty["partyid"] = new EntityReference("systemuser", new Guid(senderGuid));
            toParty["partyid"] = new EntityReference("contact", new Guid(receiverGuid));

            email["from"] = new Entity[] { fromParty };
            email["to"] = new Entity[] { toParty };
            email["subject"] = emailTemplateName;
            email["description"] = body;
            email["directioncode"] = true;
            email["regardingobjectid"] = new EntityReference("contact", new Guid(receiverGuid));

            Guid emailId = organizationService.Create(email);
            SendEmailRequest emailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            return (SendEmailResponse)organizationService.Execute(emailRequest);
        }

        // Deprecated.
        /*        private static Entity GetEmailTemplateByName(IOrganizationService service, string templateName)
                {
                    QueryExpression query = new QueryExpression("template");
                    query.ColumnSet.AddColumn("body");
                    query.Criteria.AddCondition("title", ConditionOperator.Equal, templateName);

                    EntityCollection result = service.RetrieveMultiple(query);
                    if (result.Entities.Count > 0)
                        return result.Entities[0];
                    else
                        throw new Exception($"Email template '{templateName}' not found.");
                }
        */
    }
}
