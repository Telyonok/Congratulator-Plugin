using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.ServiceModel;
using System.Text;

namespace CongratulatorPlugin
{
    public class BirthdayCongratulatorPlugin : IPlugin
    {
        private const string EmailSubject = "Birthday Congratulation";
        private const string DelayedEmailSendApiUrl = "https://prod2-28.germanywestcentral.logic.azure.com:443/workflows/aaf3fd3124f14cae92deae92697bd8da/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cc1ywoDm2OtKUyqDEEFEaKXcQ6AsnRfpjqyTAHurefM";
        private const string ContentType = "application/json";

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService organizationService = factory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.Depth > 1)  // If the plugin was called by other plugin - we ignore that call.
                return;

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
            if ((context.MessageName != "Create" && context.MessageName != "Update") ||  // Check for event type, plugin should only work for "Create" and "Update" messages.
                !context.InputParameters.Contains("Target") ||
                !(context.InputParameters["Target"] is Entity))
                return;

            Entity entity = (Entity)context.InputParameters["Target"];

            if (entity.LogicalName != "contact" ||   // Return if the entity is not Contact or has no birthdate or email specified.
                !entity.Attributes.Contains("birthdate") ||
                !entity.Attributes.Contains("emailaddress1"))   
                return;

            DateTime birthdayDate = (DateTime)entity.Attributes["birthdate"];

            tracingService.Trace("Started congratulatory email send activity."); // Log the start of operation.

            if (HasCongratulatoryEmailBeenSentThisYear(entity.Attributes["emailaddress1"].ToString(), organizationService, tracingService))
            {
                tracingService.Trace($"Congratulatory email already sent to contact {entity.Id} this year.");
                return;
            }

            if (birthdayDate.Day == DateTime.Now.Day &&
                birthdayDate.Month == DateTime.Now.Month)
            {
                tracingService.Trace("Birthday is today!");  // Log the end of operation.
                ScheduleCongratulatoryEmail(organizationService, DateTime.Now.AddMinutes(2), entity.Id, context.UserId);   // If birthday  is today - schedule it 2 minutes from now.
            }
            else
                ScheduleCongratulatoryEmail(organizationService, new DateTime(DateTime.Now.Year, birthdayDate.Month, birthdayDate.Day), entity.Id, context.UserId);

            tracingService.Trace("Ended congratulatory email send activity.");  // Log the end of operation.
        }

        private void ScheduleCongratulatoryEmail(IOrganizationService organizationService, DateTime birthdayDate, Guid receiverGuid, Guid senderGuid)
        {
            // Construct the JSON payload.
            var jsonPayload = new
            {
                scheduleDate = birthdayDate.ToString("yyyy-MM-ddTHH:mm:00Z"),
                title = EmailSubject,
                receiverGuid = receiverGuid.ToString(),
                senderGuid = senderGuid.ToString()
            };

            // Make the POST request to PA flow.
            var httpClient = new HttpClient();
            var requestUri = DelayedEmailSendApiUrl;
            var content = new StringContent(JsonConvert.SerializeObject(jsonPayload), Encoding.UTF8, ContentType);

            try
            {
                var response = httpClient.PostAsync(requestUri, content).Result;
                response.EnsureSuccessStatusCode();
            }
            catch (Exception ex)
            {
                // Log the exception
                throw new InvalidPluginExecutionException($"Failed to schedule congratulatory email: {ex.Message}", ex);
            }
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
    }
}
