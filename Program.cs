using System;
using System.Buffers;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using EAGetMail;
using System.Globalization;
using System.Diagnostics;
using System.Dynamic;

namespace Leetcode
{
    internal class Program
    {

        // Generate an unqiue email file name based on date time
        static string _generateFileName(int sequence)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.eml",
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("en-US")),
                currentDateTime.Millisecond,
                sequence);
        }

        static void Main(string[] args)
        {
            Console.WriteLine("+------------------------------------------------------------------+");
            Console.WriteLine("  Sign in with MS OAuth                                             ");
            Console.WriteLine("   If you got \"This app isn't verified\" information in Web Browser, ");
            Console.WriteLine("   click \"Advanced\" -> Go to ... to continue test.");
            Console.WriteLine("+------------------------------------------------------------------+");
            Console.WriteLine("");
            Console.WriteLine("Press any key to sign in...");
            Console.ReadKey();

            try
            {
                Program p = new Program();
                p.DoOauthAndRetrieveEmail();
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.ToString());
            }

            Console.ReadKey();
        }


        void RetrieveMailWithXOAUTH2(string userEmail, string accessToken)
        {
            try
            {
                // Create a folder named "inbox" under current directory
                // to save the email retrieved.
                string localInbox = string.Format("{0}\\inbox", Directory.GetCurrentDirectory());
                // If the folder is not existed, create it.
                if (!Directory.Exists(localInbox))
                {
                    Directory.CreateDirectory(localInbox);
                }

                // Office 365 IMAP server address
                MailServer oServer = new MailServer("outlook.office365.com",
                        userEmail,
                        accessToken, // use access token as password
                        ServerProtocol.Imap4);

                // Set IMAP OAUTH 2.0
                oServer.AuthType = ServerAuthType.AuthXOAUTH2;
                // Enable SSL/TLS connection, most modern email server require SSL/TLS by default
                oServer.SSLConnection = true;
                // Set IMAP4 SSL Port
                oServer.Port = 993;

                MailClient oClient = new MailClient("TryIt");
                // Get new email only, if you want to get all emails, please remove this line
                oClient.GetMailInfosParam.GetMailInfosOptions = GetMailInfosOptionType.NewOnly;

                Console.WriteLine("Connecting {0} ...", oServer.Server);
                oClient.Connect(oServer);

                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} email(s)\r\n", infos.Length);

                for (int i = 0; i < 5; i++)
                {
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);

                    // Receive email from email server
                    Mail oMail = oClient.GetMail(info);

                    Console.WriteLine("From: {0}", oMail.From.ToString());
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject);

                    // Generate an unqiue email file name based on date time.
                    string fileName = _generateFileName(i + 1);
                    string fullPath = string.Format("{0}\\{1}", localInbox, fileName);

                    // Save email to local disk
                    oMail.SaveAs(fullPath, true);

                    // Mark email as read to prevent retrieving this email again.
                    oClient.MarkAsRead(info, true);

                    // If you want to delete current email, please use Delete method instead of MarkAsRead
                    // oClient.Delete(info);
                }

                // Quit and expunge emails marked as deleted from server.
                oClient.Quit();
                Console.WriteLine("Completed!");
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }

        // client configuration
        // You should create your client id and client secret,
        // do not use the following client id in production environment, it is used for test purpose only.
        const string clientID = "clinte";
        const string clientSecret = "clinetSecret";
        // use IMAP scope
        const string scope = "https://outlook.office.com/IMAP.AccessAsUser.All%20offline_access%20email%20openid";
        // const string authUri = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        // const string tokenUri = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        // if your application is single tenant, please use tenant id instead of common in authUri and tokenUri
        // for example, your tenant is 669595d0-a4d7-47c5-8040-cf9970400e48, then
        const string authUri = "https://login.microsoftonline.com/tenatID/oauth2/v2.0/authorize";
        const string tokenUri = "https://login.microsoftonline.com/tenatID/oauth2/v2.0/token";

        static int GetRandomUnusedPort()
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }

        async void DoOauthAndRetrieveEmail()
        {
            // Creates a redirect URI using an available port on the loopback address.
            string redirectUri = "http://localhost:50143/";
            Console.WriteLine("redirect URI: " + redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectUri);
            Console.WriteLine("Listening ...");
            http.Start();

            // Creates the OAuth 2.0 authorization request.
            string authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}&prompt=login",
                authUri,
                scope,
                Uri.EscapeDataString(redirectUri),
                clientID
            );

            // Opens request in the browser.
            Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", authorizationRequest);

            // Waits for the OAuth authorization response.
            var context = await http.GetContextAsync();

            // Brings the Console to Focus.
            BringConsoleToFront();

            // Sends an HTTP response to the browser.
            var response = context.Response;
            string responseString = string.Format("<html><head></head><body>Please return to the app and close current window.</body></html>");
            var buffer = Encoding.UTF8.GetBytes(responseString);
            response.ContentLength64 = buffer.Length;
            var responseOutput = response.OutputStream;
            Task responseTask = responseOutput.WriteAsync(buffer, 0, buffer.Length).ContinueWith((task) =>
            {
                responseOutput.Close();
                http.Stop();
                Console.WriteLine("HTTP server stopped.");
            });

            // Checks for errors.
            if (context.Request.QueryString.Get("error") != null)
            {
                Console.WriteLine(string.Format("OAuth authorization error: {0}.", context.Request.QueryString.Get("error")));
                return;
            }

            if (context.Request.QueryString.Get("code") == null)
            {
                Console.WriteLine("Malformed authorization response. " + context.Request.QueryString);
                return;
            }

            // extracts the code
            var code = context.Request.QueryString.Get("code");
            Console.WriteLine("Authorization code: " + code);

            string responseText = await RequestAccessToken(code, redirectUri);
            Console.WriteLine(responseText);

            OAuthResponseParser parser = new OAuthResponseParser();
            parser.Load(responseText);

            var user = parser.EmailInIdToken;
            var accessToken = parser.AccessToken;

            Console.WriteLine("User: {0}", user);
            Console.WriteLine("AccessToken: {0}", accessToken);

            RetrieveMailWithXOAUTH2(user, accessToken);
        }

        async Task<string> RequestAccessToken(string code, string redirectUri)
        {
            Console.WriteLine("Exchanging code for tokens...");

            // builds the  request
            //string tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&grant_type=authorization_code",
            //    code,
            //    Uri.EscapeDataString(redirectUri),
            //    clientID
            //    );

            //  if you use it in web application, please add clientSecret parameter
            string tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&client_secret={3}&grant_type=authorization_code",
               code,
               Uri.EscapeDataString(redirectUri),
               clientID,
               clientSecret
               );


            // sends the request
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(tokenUri);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

            byte[] _byteVersion = Encoding.ASCII.GetBytes(tokenRequestBody);
            tokenRequest.ContentLength = _byteVersion.Length;

            Stream stream = tokenRequest.GetRequestStream();
            await stream.WriteAsync(_byteVersion, 0, _byteVersion.Length);
            stream.Close();

            try
            {
                // gets the response
                WebResponse tokenResponse = await tokenRequest.GetResponseAsync();
                using (StreamReader reader = new StreamReader(tokenResponse.GetResponseStream()))
                {
                    // reads response body
                    return await reader.ReadToEndAsync();
                }

            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        Console.WriteLine("HTTP: " + response.StatusCode);
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // reads response body
                            string responseText = await reader.ReadToEndAsync();
                            Console.WriteLine(responseText);
                        }
                    }
                }

                throw ex;
            }
        }

        // Hack to bring the Console window to front.

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public void BringConsoleToFront()
        {
            SetForegroundWindow(GetConsoleWindow());
        }
    }
}

