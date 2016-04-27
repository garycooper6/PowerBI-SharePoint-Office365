
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PowerBI_Office365Web.Default" EnableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        height = 800;
        width = 600;
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
                iframe = document.getElementById('powerBiFrame');
                iframe.contentWindow.postMessage(message, "*");;
            }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button Text="Sign in" OnClick="SignInButton_Click" runat="server" />
    </div>      
        <iframe id="powerBiFrame" onload="postActionLoadTile()" height="800" width="600" src="<%=PowerBIEmbedUrl %>" />        
        
        <asp:Panel ID="signinStatus" runat="server" Visible="false">
            <em>User:</em><asp:Label ID="userLabel" runat="server" /><br />
            
            <em>Access token: </em><asp:Label ID="accessTokenTextbox" runat="server" /><br />
        </asp:Panel>
    </form>
</body>
</html>
