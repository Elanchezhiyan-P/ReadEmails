using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace ReadEmails
{
    internal class Program
    {
        public static readonly string OUTLOOK_SCOPES = "https://outlook.office365.com/.default";
        public static readonly string OUTLOOK_APPID = "xxxxxxxxxxxxxxxxxx";
        public static readonly string OUTLOOK_SECRETID = "xxxxxxxxxxxx";
        public static readonly string OUTLOOK_TENANTID = "xxxxxxxxxxxxxxxxxxxx";

        private static ExchangeService? ewsClient;

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            await ReadEmails("xxxxx@xxx.com");
        }

        public static async Task<bool> ReadEmails(string emailAddress)
        {
            TimeSpan ts;
            try
            {
                await ConnectToExchangeService(emailAddress);

                if (ewsClient == null)
                {
                    Console.WriteLine("EWS Client is not initialized.");
                    return false;
                }

                string sLastEmailStamp = GetLastEmailStamp(); // Implement this method according to your needs
                if (string.IsNullOrEmpty(sLastEmailStamp) || !DateTime.TryParse(sLastEmailStamp, out DateTime datEml))
                {
                    ts = new TimeSpan(1, 0, 0, 0); // Look one day in the past if no record found
                }
                else
                {
                    ts = datEml - DateTime.Now.AddSeconds(-1);
                }

                DateTime date = DateTime.Now.Add(ts);
                SearchFilter filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);

                if (ewsClient != null)
                {
                    FindItemsResults<Item> findResults = ewsClient.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));

                    foreach (Item item in findResults)
                    {
                        EmailMessage message = EmailMessage.Bind(ewsClient, item.Id);
                        string sFromAddress = message.From.Address.ToLower();

                        // Process the email message as needed
                        Console.WriteLine($"Email from: {sFromAddress}");
                        Console.WriteLine($"Subject {message.Subject}");
                        Console.WriteLine($"Subject {message.Subject}");
                        Console.WriteLine($"Priority: {message.Importance}");
                        Console.WriteLine($"Has any attachments: {message.HasAttachments}");
                        Console.WriteLine("*********************************");
                        Console.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred in ReadEmail: " + ex.ToString());
            }
            return true;
        }

        //It asynchronously connects to an Exchange service using OAuth authentication. It retrieves an access token, sets up the ExchangeService client with the token and specified email address, and returns true if successful, or logs and throws an exception if an error occurs.
        private static async Task<bool> ConnectToExchangeService(string emailAddress)
        {
            try
            {
                string accessToken = await GetToken();
                ewsClient = new ExchangeService
                {
                    Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                    Credentials = new OAuthCredentials(accessToken),
                    ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, emailAddress)
                };
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", emailAddress);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred in ConnectToExchangeService: " + ex.ToString());
                throw;
            }
        }

        //It asynchronously retrieves an OAuth2 access token for Outlook services using client credentials (client ID, secret, tenant ID). It handles errors and returns the token if successful.
        private static async Task<string> GetToken()
        {
            string sToken = "";

            try
            {
                var cca = ConfidentialClientApplicationBuilder
                    .Create(OUTLOOK_APPID)
                    .WithClientSecret(OUTLOOK_SECRETID)
                    .WithTenantId(OUTLOOK_TENANTID)
                    .Build();

                var ewsScopes = new[] { OUTLOOK_SCOPES };
                var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();
                sToken = authResult.AccessToken;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred in getting email token: " + ex.ToString());
            }

            return sToken;
        }

        //Need to fetch it from database so that the program can know when it was last read
        private static string GetLastEmailStamp()
        {
            return null;
        }
    }
}
