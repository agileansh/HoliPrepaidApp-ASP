<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--% If Session("MM_Username")<>"Entry" Then 
	Response.Redirect("index.asp?failed=true")
End If
%-->
<!--%

' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp?failed=true"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%-->
<!--#include file="Connections/HoliCon.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script language="javascript">// When ready...
window.addEventListener("load",function() {
	// Set a timeout...
	setTimeout(function(){
		// Hide the address bar!
		window.scrollTo(0, 1);
	}, 0);
});
</script>
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="viewport" content="width=device-width; initial-scale=1; minimal-ui">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width" />
<title>Holi Prepaid</title>
<link href="styler.css" rel="stylesheet" type="text/css" />
</head>
<body >

<table align="center" border="0" width="300">
<tr>
<td>
<form method="post" action="find.asp" name="LastName">
  <div align="center"><img src="Holi Logo.jpg" width="200" height="99" align="middle" />
    <br/> 
    <Br/>
  <input type="search" placeholder="Enter Last Name" name="lname" />
  
  <input type="submit" value="Color Search">
  </div>
</form>
</td>

</tr>
</table>
</body>
</html>
