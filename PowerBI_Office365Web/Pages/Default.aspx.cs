using Microsoft.IdentityModel.Clients.ActiveDirectory;
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
        public string PowerBIEmbedUrl { get; set; }
        string baseUri = "https://api.powerbi.com/beta/myorg/";
        

        protected void Page_Load(object sender, EventArgs e)
        {
            //Test for AuthenticationResult
            if (Session["authResult"] != null)
            {
                //Get the authentication result from the session
                authResult = (AuthenticationResult)Session["authResult"];

                //Show Power BI Panel
                this.signinStatus.Visible = true;

                //Set user and token from authentication result
                userLabel.Text = authResult.UserInfo.DisplayableId;
                accessTokenTextbox.Text = authResult.AccessToken;

                if (!Page.IsPostBack)
                {
                    var dashboards = GetDashboards();
                    this.DashboardsList.DataSource = dashboards.value;
                    this.DashboardsList.DataBind();
                    this.DashboardsList.Items.Insert(0, new ListItem("---", ""));
                }

            }
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
            this.TilesList.Items.Insert(0, new ListItem("---", ""));            
        }

        protected void TilesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.PowerBIEmbedUrl = this.TilesList.SelectedValue;
        }
    }
}