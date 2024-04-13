<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/HoliCon.asp" -->

<!--% If Session("MM_Username")<>"Entry" Then 
	Response.Redirect("index.asp?failed=true")
End If
%-->
<!--%
'response.write(Session("MM_Username"))

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
'response.Write(I)
Lname=(request.QueryString("LastName"))
FName=(Request.QueryString("FirstName"))
Unique=(Request.QueryString("Unique"))
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_HoliCon_STRING
Recordset1_cmd.CommandText = "SELECT * FROM PrePaidGate WHERE LastName ='"&Lname&"' AND FirstName ='"& FName & "' AND ID="& Unique &""
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
 
Recordset1_numRows = 1
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
<title>Spanish Fork Festival of Colors</title>
<link href="styler.css" rel="stylesheet" type="text/css">
</head>

<body>

<P><span class="Heading">
  <% =(Recordset1.Fields.Item("LastName").Value)%>,     
  <% =(Recordset1.Fields.Item("FirstName").Value)%>
  
</span></p>

<span class="Text">



<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_HoliCon_STRING
    MM_editCmd.CommandText = "UPDATE PrePaidGate SET LastName = ?, FirstName = ?, Entries = ?, Checked = ? WHERE ID = ?" 
    MM_editCmd.Prepared = true
    ' ID Update is Removed
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("LastName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("FirstName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 202, 1, 255, Request.Form("Entries")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 202, 1, 255, Request.Form("Checked")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.querystring("Unique") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("Unique")

End If
%>
<%


Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_HoliCon_STRING
Recordset1_cmd.CommandText = "SELECT * FROM PrePaidGate WHERE ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<% Dim Checked
 	Checked=(Recordset1.Fields.Item("Checked").Value)
	'response.Write(checked)
	If Checked="1" Then response.Write("<style=CSS class=text_ckd>")
	%>
<% 	'Entries=(Recordset1.Fields.Item("Entries").Value)
	'response.write("ENTERED")
	'ID=(Recordset1.Fields.Item("ID").Value)
%>
	    
    <% If (Recordset1.Fields.Item("Entries").Value)="Multiple"  Then 
	Response.Write("Entries:<span id=warning> " & (Recordset1.Fields.Item("Entries").Value) &"</span></br>")
	Else Response.Write("Entries: Single")
	End If %>


<% If Checked="1" Then response.Write("<style=CSS class=text_ckd>")
	%>
</span>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  
  <span class="Text">
    <input type="hidden" name="LastName" value="<%=(Recordset1.Fields.Item("LastName").Value)%>" size="32" />
    <input type="hidden" name="FirstName" value="<%=(Recordset1.Fields.Item("FirstName").Value)%>" size="32" />
    
  
    <input type="hidden" name="Entries" value="<%=(Recordset1.Fields.Item("Entries").Value)%>" size="32" />
   <input type="hidden" name="Checked" value="1" size="32" />
    <br />
    <Br />
    <Br />
    <Br />
    </span>
  <span class="Heading">
  <% If Checked="1" Then 
	  	response.Write("<a href=index_entry.asp>New Search</a>")
Else 
		response.Write("<input type=submit value=Checkout!>")
End if%>
  </span>  <span class="Text">
    
    <input type="hidden" name="Unique" value="<% =Recordset1.Fields.Item("ID")%>" />
    <input type="hidden" name="MM_update" value="form1" />
    <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>" />
    </span>
</form>
</span>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
