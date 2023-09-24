using System.Text;
using System.Net;
using Newtonsoft.Json;
using Microsoft.Exchange.WebServices.Data;
using System.IO.Compression;
using System.Xml;

namespace ClassificationBackend
{


    class ClassificationRequest
    {
        public string token;
        public string ews_url;
        public string[] attachment_ids;
    }

    class ClassificationResponse
    {
        public string classification;
        public string error;
    }

    class HttpServer
    {
        public static HttpListener listener;
        public static string pageData =
            "<!DOCTYPE html>" +
            "<html>" +
            "  <head>" +
            "    <title>Invalid request</title>" +
            "  </head>" +
            "  <body>" +
            "    <p>API server got invalid request</p>" +
            "  </body>" +
            "</html>";

        internal enum Classification
        {
            None,
            White,
            Green,
            Amber,
            Red
        }

        private static readonly Dictionary<Classification, string> classificationProperties = new Dictionary<Classification, string> {
            { Classification.White, "TLP:WHITE" },
            { Classification.Green, "TLP:GREEN" },
            { Classification.Amber, "TLP:AMBER" },
            { Classification.Red, "TLP:RED" },
       };

        public static async System.Threading.Tasks.Task HandleIncomingConnections()
        {

            while (true)
            {
                // Will wait here until we hear from a connection
                HttpListenerContext ctx = listener.GetContext();

                // TODO: simplification, thread <-> request, will not scale properly
                // use a worker pool instead
                new Thread(async () =>
                {
                Thread.CurrentThread.IsBackground = true;
                
                HttpListenerRequest req = ctx.Request;
                HttpListenerResponse resp = ctx.Response;

                // Print out some info about the request
                Console.WriteLine("Request for {0}", req.Url.ToString());
                Console.WriteLine(req.HttpMethod);
                Console.WriteLine(req.UserHostName);
                Console.WriteLine(req.UserAgent);
                Console.WriteLine();

             
                if (req.Url.AbsolutePath != "/api/attachment")
                {
                    Console.WriteLine("Invalid request");

                    // Write the response info
                    byte[] pageBytes = Encoding.UTF8.GetBytes(pageData);
                    resp.ContentType = "text/html";
                    resp.ContentEncoding = Encoding.UTF8;
                    resp.ContentLength64 = pageBytes.LongLength;

                    // Write out to the response stream (asynchronously), then close it
                    await resp.OutputStream.WriteAsync(pageBytes, 0, pageBytes.Length);
                    resp.Close();
                    return;
                }

                if (req.HttpMethod == "OPTIONS")
                {
                    // CORS preflight
                    resp.Headers.Clear();
                    resp.SendChunked = false;
                    resp.StatusCode = 204; // No Content
                    resp.AddHeader("Access-Control-Allow-Origin", "*");
                    resp.AddHeader("Access-Control-Allow-Methods", "POST");
                    resp.AddHeader("Access-Control-Allow-Private-Network", "true");
                    resp.AddHeader("Access-Control-Allow-Headers", "*");
                    resp.AddHeader("Access-Control-Max-Age", "86400");
                    resp.Close();
                    return;
                }

                if (req.HttpMethod == "POST")
                {

                    // Potential client request, try to read JSON

                    // TODO limit request size, catch exceptions and return error to client
                    string reqData = GetRequestPostData(req);
                    Console.WriteLine(reqData);
                    ClassificationRequest c = JsonConvert.DeserializeObject<ClassificationRequest>(reqData);

                    Classification output = ProcessClassificationRequestEWS(c);

                    //  Do classification
                    ClassificationResponse cResp = new ClassificationResponse();
                    cResp.classification = classificationProperties[output];
                    cResp.error = ""; // TODO fill this

                    string json = JsonConvert.SerializeObject(cResp);
                    // Write out to the response stream (asynchronously), then close it

                    byte[] data = Encoding.UTF8.GetBytes(json);

                    resp.AddHeader("Access-Control-Allow-Origin", "*");
                    resp.ContentType = "application/json";
                    resp.ContentEncoding = Encoding.UTF8;
                    resp.ContentLength64 = data.LongLength;

                    // Write out to the response stream (asynchronously), then close it
                    resp.OutputStream.Write(data, 0, data.Length);
                    resp.Close();
                }


                  }).Start(); // Do not track to save resources
            }
        }

        public static void Main(string[] args)
        {
            // HTTPS Cert GUID: {85718C98-BE2E-4FC2-A040-562EFB2953EE}
            // Note this is not associated to any real registry, but identifies certificate bindings

            // Create self-signed cert:
            // powershell -Command 'New-SelfSignedCertificate -DnsName localhost,exchange.classification.lab -CertStoreLocation cert:\LocalMachine\My -NotAfter (Get-Date).AddYears(10)'
            // netsh http add sslcert ipport=0.0.0.0:4430 certhash=<THUMBPRINT> appid={85718C98-BE2E-4FC2-A040-562EFB2953EE}
            // netsh http add urlacl url=https://+:4430/ user=Everyone

            // Create a Http server and start listening for incoming connections
            listener = new HttpListener();
            listener.Prefixes.Add("https://+:4430/");
            //listener.Prefixes.Add("https://exchange.classification.lab:4430/");

            listener.Start();
            Console.WriteLine("Listening for connections on port 4430");

            // Handle requests
            System.Threading.Tasks.Task listenTask = HandleIncomingConnections();
            listenTask.GetAwaiter().GetResult();

            // Close the listener
            listener.Close();
        }

        public static string GetRequestPostData(HttpListenerRequest request)
        {
            if (!request.HasEntityBody)
            {
                return "";
            }
            System.IO.Stream body = request.InputStream;
            System.IO.StreamReader reader = new System.IO.StreamReader(body, request.ContentEncoding);
            return reader.ReadToEnd();
        }

        // Inspired by https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item#use-the-ews-managed-api-to-get-the-attachments
        public static Classification ProcessClassificationRequestEWS(ClassificationRequest request)
        {
            var attachmentNames = new List<string>();
            Classification highestClassification = Classification.White;

            // Create an ExchangeService object, set the credentials and the EWS URL.
            ExchangeService service = new ExchangeService();
            service.Credentials = new OAuthCredentials(request.token);
            service.Url = new Uri(request.ews_url);

            var attachmentIds = new List<string>();

            foreach (string id in request.attachment_ids)
            {
                attachmentIds.Add(id);
            }

            // Call the GetAttachments method to retrieve the attachments on the message.
            // This method results in a GetAttachments EWS SOAP request and response
            // from the Exchange server.
            var getAttachmentsResponse =
              service.GetAttachments(attachmentIds.ToArray(),
                                      null,
                                      new PropertySet(BasePropertySet.FirstClassProperties,
                                                      ItemSchema.MimeContent));

            if (getAttachmentsResponse.OverallResult != ServiceResult.Success)
            {
                // TODO handle error out of band, inform user
                Console.WriteLine("Service request failed!");
                return Classification.None;
            }

            foreach (var attachmentResponse in getAttachmentsResponse)
            {
                attachmentNames.Add(attachmentResponse.Attachment.Name);

                // Write the content of each attachment to a stream.
                if (attachmentResponse.Attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
                    Stream s = new MemoryStream(fileAttachment.Content);
                    Console.WriteLine("Reading file attachment {0}", fileAttachment.FileName);

                    // Process the contents of the attachment here.
                    // TODO: Read zip, extract custom XML, determine classification

                    ZipArchive archive = new ZipArchive(s);
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith("custom.xml", StringComparison.Ordinal))
                        {
                            Stream propsXml = entry.Open(); // .Open will return a stream

                            XmlDocument doc = new XmlDocument();
                            doc.Load(propsXml);
                            XmlNodeList properties = doc.GetElementsByTagName("property");
                            foreach (XmlNode node in properties)
                            {
                                Console.WriteLine(node.Name + ": " + node.FirstChild.InnerText);
                                Classification key = classificationProperties.FirstOrDefault(x => x.Value == node.FirstChild.InnerText).Key;
                                if (key > highestClassification)
                                {
                                    highestClassification = key;
                                }
                            }
                            break;
                        }
                    }

                }

                /*if (attachmentResponse.Attachment is ItemAttachment)
                {
                    ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
                    Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
                    // Process the contents of the attachment here.
                }*/

            }


            // Return the names and number of attachments processed for display
            // in the add-in UI.
            //var response = new AttachmentSampleServiceResponse();
            //response.attachmentNames = attachmentNames.ToArray();
            //response.attachmentsProcessed = attachmentsProcessedCount;

            return highestClassification;
        }

    }

}