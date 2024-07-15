using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.ServiceModel;
using System.Xml;

namespace CongratulatorPlugin
{
    public class BirthdayCongratulatorPlugin : IPlugin
    {
        private const string EmailSubject = "Happy Birthday";
        private readonly string emailTemplateXML = string.Empty;
        public BirthdayCongratulatorPlugin(string unSecureConfig, string secureConfig)
        {
            emailTemplateXML = unSecureConfig;
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService organizationService = factory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                InitiateCongratulationActivity(context, organizationService, tracingService);
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

        private void InitiateCongratulationActivity(IPluginExecutionContext context, IOrganizationService organizationService, ITracingService tracingService)
        {
            if (context.MessageName != "Create" ||  // Check for event type, plugin should only work for "Create" messages.
                !context.InputParameters.Contains("Target") ||
                !(context.InputParameters["Target"] is Entity))
                return;

            Entity entity = (Entity)context.InputParameters["Target"];

            if (entity.LogicalName != "contact" ||   // Return if the entity is not Contact or has no birthdate or email specified.
                !entity.Attributes.Contains("birthdate") ||
                !entity.Attributes.Contains("emailaddress1"))   
                return;

            DateTime birthdayDate = (DateTime)entity.Attributes["birthdate"];

            if (birthdayDate.Day != DateTime.Now.Day || birthdayDate.Month != DateTime.Now.Month) // Return if today is not Contact's birthday.
                return;

            tracingService.Trace("Started congratulatory email send activity."); // Log the start of operation.

            if (HasCongratulatoryEmailBeenSentThisYear(entity.Attributes["emailaddress1"].ToString(), organizationService, tracingService))
            {
                tracingService.Trace($"Congratulatory email already sent to contact {entity.Id} this year.");
                return;
            }

            SendEmailResponse response = SendCongratulatoryEmail(entity, context.UserId, organizationService, tracingService);
            tracingService.Trace("Email response: " + response.Results.Values.First().ToString());
            tracingService.Trace("Ended congratulatory email send activity.");  // Log the end of operation.
        }

        private bool HasCongratulatoryEmailBeenSentThisYear(string contactEmail, IOrganizationService organizationService, ITracingService tracingService)
        {
            string fetchXml = $@"
            <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                <entity name=""email"">
                    <attribute name=""subject""/>
                    <order attribute=""subject"" descending=""false""/>
                    <filter type=""and"">
                        <condition attribute=""createdon"" operator=""this-year""/>
                        <condition attribute=""torecipients"" operator=""like"" value=""%{contactEmail}%""/>
                        <condition attribute=""subject"" operator=""like"" value=""{EmailSubject}%""/>
                    </filter>
                </entity>
            </fetch>";  // Get all emails sent to this contact this year with Happy Birthday subject.

            var emails = organizationService.RetrieveMultiple(new FetchExpression(fetchXml)).Entities;
            bool emailSentThisYear = emails.Count > 0; // If there is already one sent - no need to send another one.

            return emailSentThisYear;
        }

        private SendEmailResponse SendCongratulatoryEmail(Entity entity, Guid senderGuid, IOrganizationService organizationService, ITracingService tracingService)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(emailTemplateXML);
            Entity email = new Entity("email");

            Entity fromParty = new Entity("activityparty");
            Entity toParty = new Entity("activityparty");

            fromParty["partyid"] = new EntityReference("systemuser", senderGuid);
            toParty["partyid"] = new EntityReference("contact", entity.Id);

            email["from"] = new Entity[] { fromParty };
            email["to"] = new Entity[] { toParty };
            email["subject"] = EmailSubject;
            email["description"] = string.Format(doc.SelectSingleNode("//body").InnerText, entity.Attributes["fullname"]); // Take template from plugin config.
            email["directioncode"] = true;
            email["regardingobjectid"] = new EntityReference("contact", entity.Id);

            Guid emailId = organizationService.Create(email);
            SendEmailRequest emailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            return (SendEmailResponse)organizationService.Execute(emailRequest);
        }
    }
}
