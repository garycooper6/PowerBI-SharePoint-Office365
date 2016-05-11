
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PowerBI_Office365Web.Default" EnableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        window.onload = function () {
            // listen for message to receive tile click messages.
            if (window.addEventListener) {
                window.addEventListener("message", receiveMessage, false);
            } else {
            window.attachEvent("onmessage", receiveMessage);
            }
        
        }
        function receiveMessage(event) {
            if (event.data) {
                try {
                    messageData = JSON.parse(event.data);
                    if (messageData.event === "tileClicked") {
                        //Get IFrame source and construct dashboard url
                        iFrameSrc = document.getElementById(event.srcElement.iframe.id).src;

                        //Split IFrame source to get dashboard id
                        var dashboardId = iFrameSrc.split("dashboardId=")[1].split("&")[0];

                        //Get PowerBI service url
                        urlVal = iFrameSrc.split("/embed")[0] + "/dashboards/{0}";
                        urlVal = urlVal.replace("{0}", dashboardId);

                        window.open(urlVal);
                    }
                }
                catch (e) {
                    // In a production app, handle exception
                }
            }
        }
        height = 800;
        width = 1400;
        // Post the authentication token to the IFrame.
        function postActionLoadTile() {
                
            // get the access token.
            //accessToken = document.getElementById('MainContent_accessTokenTextbox').value;
            var accessToken = '<%= authResult == null ? "" : authResult.AccessToken%> ';
            // return if no a
            if ("" === accessToken) {
                console.log("NO ACCESS TOKEN");
                return;
            }               

            var h = height;
            var w = width;

            // construct the post message structure
            var m = { action: "loadTile", accessToken: accessToken, height: h, width: w };
            message = JSON.stringify(m);

            // push the message.
            iframe = document.getElementById('powerBiTileFrame');
            iframe.contentWindow.postMessage(message, "*");;
        }

        function postActionLoadReport() {

            // get the access token.
            var accessToken = '<%= authResult == null ? "" : authResult.AccessToken%> ';

            // return if no a
            if ("" === accessToken) {
                console.log("NO ACCESS TOKEN");
                return;
            }

            // construct the push message structure
            // this structure also supports setting the reportId, groupId, height, and width.
            // when using a report in a group, you must provide the groupId on the iFrame SRC
            var m = { action: "loadReport", accessToken: accessToken };
            message = JSON.stringify(m);

            // push the message.
            iframe = document.getElementById('powerBiReportFrame');
            iframe.contentWindow.postMessage(message, "*");;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button Text="Sign in" OnClick="SignInButton_Click" runat="server" ID="SignInButton" />
    </div>      
        <em>Dashboards</em>
        <asp:DropDownList runat="server" ID="DashboardsList" AutoPostBack="true" OnSelectedIndexChanged="DashboardsList_SelectedIndexChanged" DataTextField="displayName" DataValueField="id" />
        <br />
        <em>Tiles</em>
        <asp:DropDownList runat="server" ID="TilesList" AutoPostBack="true" OnSelectedIndexChanged="TilesList_SelectedIndexChanged" DataTextField="title" DataValueField="embedUrl" />
        <hr />
        <em>Or, select a report</em>
        <asp:DropDownList runat="server" ID="ReportsList" AutoPostBack="true" OnSelectedIndexChanged="ReportsList_SelectedIndexChanged" DataTextField="name" DataValueField="embedUrl" />
        <br />
        <asp:Button ID="EmbedInSPButton" runat="server" Text="Embed in SharePoint" OnClick="EmbedInSPButton_Click" />
        <asp:Panel ID="BiTileFramePanel" runat="server" Visible="false">
            <iframe id="powerBiTileFrame" onload="postActionLoadTile()" height="800" width="1400" src="<%=PowerBITileEmbedUrl %>" /> 
        </asp:Panel>
        <asp:Panel ID="BiReportFramePanel" runat="server" Visible="false">
            <iframe id="powerBiReportFrame" onload="postActionLoadReport()" height="800" width="1400" src="<%=PowerBIReportEmbedUrl %>" />
        </asp:Panel>
        
        
        <asp:Label ID="lblStatus" runat="server" />
            
        <asp:Panel ID="signinStatus" runat="server" Visible="false">
            <em>User:</em><asp:Label ID="userLabel" runat="server" /><br />
            
            <em>Access token: </em><asp:Label ID="accessTokenTextbox" runat="server" /><br />

            
        </asp:Panel>
    </form>
</body>
</html>
