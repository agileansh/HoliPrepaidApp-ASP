<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/HoliCon.asp" -->

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
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("lname") <> "") Then 
  Recordset1__MMColParam = Request.Form("lname")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_HoliCon_STRING
Recordset1_cmd.CommandText = "SELECT * FROM PrePaidGate WHERE LastName LIKE ? ORDER BY LastName ASC" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 255, Recordset1__MMColParam &"%") ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
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
<meta name="viewport" content="width=device-width; initial-scale=1; minimal-ui" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width" />
<title>Untitled Document</title>
<link href="styler.css" rel="stylesheet" type="text/css" />
</head>

<body ><table width="200" border="0">
  <tr>
    
  
    <td>

<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
    <span class="Text"><a href="display_entry.asp?LastName=<%=(Recordset1.Fields.Item("LastName").Value)%>&FirstName=<%=(Recordset1.Fields.Item("FirstName").Value)%>&Unique=<%=(Recordset1.Fields.Item("ID").Value)%>">
<% =(Recordset1.Fields.Item("LastName").Value) & ", " %> <% =(Recordset1.Fields.Item("FirstName").Value)%></br></br>
    </a>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</span>
<% If Recordset1.EOF And Recordset1.BOF Then %>
    <span class="Text">No Record Found</span><br />
    <br />
    <Form action="index_entry.asp"><input type="submit" value="New Search" /></Form>
  <% End If ' end Recordset1.EOF And Recordset1.BOF %>
  </td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
