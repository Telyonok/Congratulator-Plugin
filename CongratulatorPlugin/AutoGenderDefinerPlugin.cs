using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System;
using System.Xml;

namespace CongratulatorPlugin
{
    public class AutoGenderDefinerPlugin : IPlugin
    {
        public const int ISO3166GermanyId = 276;
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
                InitiateGenderDefinerActivity(context, organizationService, tracingService);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("AutoGenderDefinerPlugin: {0}.", ex.Message);  // Log exception.
                throw;
            }
        }
        private void InitiateGenderDefinerActivity(IPluginExecutionContext context, IOrganizationService organizationService, ITracingService tracingService)
        {
            if ((context.MessageName != "Create" && context.MessageName != "Update") ||  // Check for event type, plugin should only work for "Create" and "Update" messages.
                !context.InputParameters.Contains("Target") ||
                !(context.InputParameters["Target"] is Entity))
                return;
            Entity entity = (Entity)context.InputParameters["Target"];

            if (entity.LogicalName != "contact" ||   // Return if the entity is not Contact or has no fullname specified.
                !entity.Attributes.Contains("fullname"))
                return;

            tracingService.Trace("Started gender definer activity."); // Log the start of operation.
            string response = QNCWebServiceClient.UCheckName(276, (string)entity.Attributes["fullname"]);

            tracingService.Trace("API request successful.");

            tracingService.Trace("Extracting SexCode.");
            int newGenderCode = ExtractSexCode(response);
            tracingService.Trace($"New GenderCode = {newGenderCode}.");
            entity.Attributes["gendercode"] = new OptionSetValue(newGenderCode);
            organizationService.Update(entity);
            tracingService.Trace("Ended gender definer activity."); // Log the start of operation.
        }

        private static int ExtractSexCode(string response)
        {
            // Load the XML response into an XmlDocument
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);

            // Find the SexCode element
            XmlNode sexCodeNode = doc.GetElementsByTagName("SexCode")[0];

            // Extract the SexCode value
            if (sexCodeNode != null)
            {
                string sexCode = sexCodeNode.InnerText;
                if (sexCode == "MALE")
                    return 1;
                else if (sexCode == "FEMALE")
                    return 2;
                else
                    throw new Exception("SexCode could not be defined.");
            }
            else
            {
                throw new Exception("SexCode was null.");
            }
        }
    }
}
