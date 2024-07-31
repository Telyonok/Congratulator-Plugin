using System.Net;
using System.Xml;

namespace CongratulatorPlugin
{
    public static class QNCWebServiceClient
    {
        private const string _url = "https://qws2.de/QWS/QNCWebService.asmx";

        public static string UCheckName(int countryId, string fullname)
        {
            // Create a new WebClient instance.
            WebClient client = new WebClient();

            // Create a new XmlDocument instance.
            XmlDocument document = new XmlDocument();

            // Create the SOAP request.
            string request = CreateSoapRequest(countryId, fullname);

            // Set the content type of the request.
            client.Headers["Content-Type"] = "text/xml; charset=utf-8";

            // Set the SOAP action.
            client.Headers["SOAPAction"] = "http://www.qaddress.de/webservices/UCheckName";

            // Send the request and get the response.
            string response = client.UploadString(_url, request);

            return response;
        }

        private static string CreateSoapRequest(int countryId, string fullname)
        {
            // Create the SOAP request.
            string request = @"<?xml version=""1.0"" encoding=""utf-8""?>
            <soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
              <soap:Body>
                <UCheckName xmlns=""http://www.qaddress.de/webservices"">
                  <UserName>string</UserName>
                  <UserPassword>string</UserPassword>
                  <SourceName>
                    <m_iCountryID>{0}</m_iCountryID>
                    <m_szName1>{1}</m_szName1>
                  </SourceName>
                </UCheckName>
              </soap:Body>
            </soap:Envelope>";

            // Replace the placeholder values with your actual values.
            request = string.Format(request, countryId, fullname);

            return request;
        }
    }
}
