using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PowerBI_Office365Web.Models;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PowerBI_Office365Web
{
    public partial class Default : System.Web.UI.Page
    {
        public AuthenticationResult authResult { get; set; }
        public string PowerBITileEmbedUrl { get; set; }
        public string PowerBIReportEmbedUrl { get; set; }
        string baseUri = "https://api.powerbi.com/beta/myorg/";

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while verifying your SharePoint Context. Probably the required query string parameters are gone because you've logged in. Please click back once or twice and refresh the page when the query string parameters are available.");
                    Response.End();
                    break;
            }
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            //Test for AuthenticationResult
            if (Session["authResult"] != null)
            {
                //Get the authentication result from the session
                authResult = (AuthenticationResult)Session["authResult"];

                //Show Power BI Panel and hide button
                this.signinStatus.Visible = true;
                this.SignInButton.Visible = false;

                //Set user and token from authentication result
                userLabel.Text = authResult.UserInfo.DisplayableId;
                accessTokenTextbox.Text = authResult.AccessToken;

                if (!Page.IsPostBack)
                {
                    var dashboards = GetDashboards();
                    this.DashboardsList.DataSource = dashboards.value;
                    this.DashboardsList.DataBind();
                    this.DashboardsList.Items.Insert(0, new System.Web.UI.WebControls.ListItem("---", ""));

                    var reports = GetReports();
                    this.ReportsList.DataSource = reports.value;
                    this.ReportsList.DataBind();
                    this.ReportsList.Items.Insert(0, new System.Web.UI.WebControls.ListItem("---", ""));

                }                
            }
        }

        private ReportsResponse GetReports()
        {
            var url = "https://api.powerbi.com/beta/myorg/reports";
            var response = GetString(url);
            return JsonConvert.DeserializeObject<ReportsResponse>(response);
        }

        private DashboardsResponse GetDashboards()
        {
            var url = "https://api.powerbi.com/beta/myorg/dashboards";
            var response = GetString(url);
            return JsonConvert.DeserializeObject<DashboardsResponse>(response);
        }

        private DashboardTileResponse GetTiles(string dashboardId)
        {
            var url = string.Format("https://api.powerbi.com/beta/myorg/dashboards/{0}/tiles", dashboardId);
            var response = GetString(url);
            return JsonConvert.DeserializeObject<DashboardTileResponse>(response);
        }

        private string GetString(string url)
        {
            var webRequest = (HttpWebRequest) WebRequest.Create(url);
            webRequest.Headers.Add("Authorization", String.Format("Bearer {0}", authResult.AccessToken));
            webRequest.Method = "GET";
            webRequest.ContentLength = 0;
            var response = webRequest.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            return responseString;

        }
        protected void SignInButton_Click(object sender, EventArgs e)
        {
            //Create a query string
            //Create a sign-in NameValueCollection for query string
            var @params = new NameValueCollection
            {
                //Azure AD will return an authorization code.
                //See the Redirect class to see how "code" is used to AcquireTokenByAuthorizationCode
                {"response_type", "code"},

                //Client ID is used by the application to identify themselves to the users that they are requesting permissions from.
                //You get the client id when you register your Azure app.
                {"client_id", Properties.Settings.Default.ClientID},

                //Resource uri to the Power BI resource to be authorized
                {"resource", "https://analysis.windows.net/powerbi/api"},

                //After user authenticates, Azure AD will redirect back to the web app
                {"redirect_uri", "https://localhost/PowerBI_Office365Web/Pages/Redirect.aspx"}
            };

            //Create sign-in query string
            var queryString = HttpUtility.ParseQueryString(string.Empty);
            queryString.Add(@params);

            //Redirect authority
            //Authority Uri is an Azure resource that takes a client id to get an Access token
            string authorityUri = "https://login.windows.net/common/oauth2/authorize/";
            Response.Redirect(String.Format("{0}?{1}", authorityUri, queryString));
        }

        protected void DashboardsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            var tiles = GetTiles(this.DashboardsList.SelectedValue);
            this.TilesList.DataSource = tiles.value;
            this.TilesList.DataBind();
            this.TilesList.Items.Insert(0, new System.Web.UI.WebControls.ListItem("---", ""));            
        }

        protected void TilesList_SelectedIndexChanged(object sender, EventArgs e)
        {         
            this.PowerBITileEmbedUrl = this.TilesList.SelectedValue;
            this.BiTileFramePanel.Visible = true;
        }

        protected void ReportsList_SelectedIndexChanged(object sender, EventArgs e)
        {            
            this.PowerBIReportEmbedUrl = this.ReportsList.SelectedValue;
            this.BiReportFramePanel.Visible = true;
        }
        static byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        protected void EmbedInSPButton_Click(object sender, EventArgs e)
        {
            if (authResult == null || string.IsNullOrEmpty(this.ReportsList.SelectedValue))
            {
                return;
            }
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                var folder = clientContext.Site.RootWeb.Lists.GetByTitle("Web Part Gallery").RootFolder;
                clientContext.Load(folder);
                clientContext.ExecuteQuery();

                // first, get the content of the web part definition
                var webPartDefininition = System.IO.File.ReadAllText(Server.MapPath("..\\powerbiwebpart.webpart"));

                // next, populate a variable with the web part content
                var content = 
                @"<script>function postActionLoadReport() {
                    // get the access token.
                    var accessToken = '" + authResult.AccessToken  + @"';

                    // return if no a
                    if ("""" === accessToken) {
                        console.log(""NO ACCESS TOKEN"");
                        return;
                    }

                    // construct the push message structure
                    // this structure also supports setting the reportId, groupId, height, and width.
                    // when using a report in a group, you must provide the groupId on the iFrame SRC
                    var m = { action: ""loadReport"", accessToken: accessToken };
                    message = JSON.stringify(m);

                    // push the message.
                    iframe = document.getElementById('powerBiReportFrame');
                    iframe.contentWindow.postMessage(message, ""*"");;
                }</script>
                <iframe id=""powerBiReportFrame"" onload=""postActionLoadReport()"" 
                    height=""800"" width=""1400"" src=""" + this.ReportsList.SelectedValue + @""" />";

                // put the web part content into the definition
                webPartDefininition = webPartDefininition.Replace("{{TOKEN_CONTENT}}", $"<![CDATA[{content}]]>");
                
                // and upload it
                FileCreationInformation fileInfo = new FileCreationInformation();
                    
                fileInfo.Content = GetBytes(webPartDefininition);
                fileInfo.Overwrite = true;
                fileInfo.Url = "userprofileinformation.webpart";
                Microsoft.SharePoint.Client.File file = folder.Files.Add(fileInfo);
                clientContext.ExecuteQuery();
                

                // Let's update the group for just uploaded web part
                var list = clientContext.Site.RootWeb.Lists.GetByTitle("Web Part Gallery");
                CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(100);
                Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(camlQuery);
                clientContext.Load(items);
                clientContext.ExecuteQuery();
                foreach (var item in items)
                {
                    // Just random group name to differentiate it from the rest
                    if (item["FileLeafRef"].ToString().ToLowerInvariant() == "userprofileinformation.webpart")
                    {
                        item["Group"] = "Add-in Script Part";
                        item.Update();
                        clientContext.ExecuteQuery();
                    }
                }

                lblStatus.Text = string.Format("Add-in script part has been added to web part gallery. You can find 'User Profile Information' script part under 'Add-in Script Part' group in the <a href='{0}'>host web</a>.", spContext.SPHostUrl.ToString());
            }
        }
    }
}