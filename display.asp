<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/HoliCon.asp" -->
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
Recordset1_cmd.CommandText = "SELECT * FROM SpanishFork2015 WHERE LastName ='"&Lname&"' AND FirstName ='"& FName & "' AND ID="& Unique &""
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
    MM_editCmd.CommandText = "UPDATE SpanishFork2015 SET LastName = ?, FirstName = ?, GuestLast = ?, GuestFirst = ?, Colors = ?, TshirtsHoli = ?, HoddieK = ?, Ganesh = ?, OmDlx = ?, Meal = ?, Bandana = ?, Buff = ?, DustMask = ?, Gita = ?, SunGlasses = ?, FreeHugs = ?, OM = ?, Entries = ?, SuperPack = ?, Checked = ? WHERE ID = ?" 
    MM_editCmd.Prepared = true
    ' ID Update is Removed
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("LastName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("FirstName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 255, Request.Form("GuestLast")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 255, Request.Form("GuestFirst")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 255, Request.Form("Colors")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 255, Request.Form("TshirtsHoli")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 202, 1, 255, Request.Form("HoddieK")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 202, 1, 255, Request.Form("Ganesh")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 202, 1, 255, Request.Form("OmDlx")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 202, 1, 255, Request.Form("Meal")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 202, 1, 255, Request.Form("Bandana")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 202, 1, 255, Request.Form("Buff")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 202, 1, 255, Request.Form("DustMask")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 202, 1, 255, Request.Form("Gita")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 202, 1, 255, Request.Form("SunGlasses")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 202, 1, 255, Request.Form("FreeHugs")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 202, 1, 255, Request.Form("OM")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param19", 202, 1, 255, Request.Form("Entries")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 202, 1, 255, Request.Form("SuperPack")) ' adVarWChar
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
Recordset1_cmd.CommandText = "SELECT * FROM SpanishFork2015 WHERE ID = ?" 
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
<% 	Colors =(Recordset1.Fields.Item("Colors").Value)
	TShirtH=(Recordset1.Fields.Item("TshirtsHoli").Value)
	TShirtG=(Recordset1.Fields.Item("Ganesh").Value)
	TShirtK=(Recordset1.Fields.Item("HoddieK").Value)
	TShirtOM=(Recordset1.Fields.Item("OM").Value)
	TShirtFH=(Recordset1.Fields.Item("FreeHugs").Value)
	TShirtOmDlx=(Recordset1.Fields.Item("OmDlx").Value)
	Bandana=(Recordset1.Fields.Item("Bandana").Value)
	Buff=(Recordset1.Fields.Item("Buff").Value)
	Gita=(Recordset1.Fields.Item("Gita").Value)
	Mask=(Recordset1.Fields.Item("DustMask").Value)
	SunGlasses=(Recordset1.Fields.Item("SunGlasses").Value)
	Meal=(Recordset1.Fields.Item("Meal").Value)
	Super=(Recordset1.Fields.Item("SuperPack").Value)
	Entries=(Recordset1.Fields.Item("Entries").Value)
	ID=(Recordset1.Fields.Item("ID").Value)
%>
	<% If Colors<>""  Then Response.Write("Color Bags: " & Colors &"</br>")%>
    <% If TShirtH<>""  Then Response.Write("Color Festival T-Shirt: " & TShirtH&" </br>")%>
    <% If TShirtG<>""  Then Response.Write("Ganesh T-Shirt: " & TShirtG &"</br>")%>
    <% If TShirtK<>""  Then Response.Write("Meditation Beeds: " & TShirtK &"</br>")%>
    <% If TShirtFH<>""  Then Response.Write("Free Hugs T-Shirt: " & TShirtFH &"</br>")%>
    <% If TShirtOM<>""  Then Response.Write("OM White T-Shirt: " & TShirtOM &"</br>")%>
    <% If TShirtOmDlx<>""  Then Response.Write("OM Black T-Shirt: " & TShirtOmDlx &"</br>")%>
    <% If Bandana<>""  Then Response.Write("Bandana: " & Bandana &"</br>")%>
    <% If Buff<>"" Then Response.Write("Buff: " & Buff &"</br>")%>
    <% If Gita<>""  Then Response.Write("Bhagwad Gita Book: " & Gita &"</br>")%>
    <% If Mask<>""  Then Response.Write("Dust Masks: " & Mask &"</br>")%>
    <% If SunGlasses<>""  Then Response.Write("Sun Glasses: " & SunGlasses &"</br>")%>
    <% If Meal<>""  Then Response.Write("Meal Coupon: " & Meal &"</br>")%>
    <% If Entries="Multiple"  Then Response.Write("Entries: " & Entries &"</br>")%>


<% If Checked="1" Then response.Write("<style=CSS class=text_ckd>")
	%>
</span>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  
  <span class="Text">
    <input type="hidden" name="LastName" value="<%=(Recordset1.Fields.Item("LastName").Value)%>" size="32" />
    <input type="hidden" name="FirstName" value="<%=(Recordset1.Fields.Item("FirstName").Value)%>" size="32" />
    
    <input type="hidden" name="GuestLast" value="<%=(Recordset1.Fields.Item("GuestLast").Value)%>" size="32" />
    
    <input type="hidden" name="GuestFirst" value="<%=(Recordset1.Fields.Item("GuestFirst").Value)%>" size="32" />
    <input type="hidden" name="Colors" value="<%=(Recordset1.Fields.Item("Colors").Value)%>" size="32" />
    
    <input type="hidden" name="TshirtsHoli" value="<%=(Recordset1.Fields.Item("TshirtsHoli").Value)%>" size="32" />
    
    <input type="hidden" name="HoddieK" value="<%=(Recordset1.Fields.Item("HoddieK").Value)%>" size="32" />
    
    <input type="hidden" name="Ganesh" value="<%=(Recordset1.Fields.Item("Ganesh").Value)%>" size="32" />
    
    <input type="hidden" name="OmDlx" value="<%=(Recordset1.Fields.Item("OmDlx").Value)%>" size="32" />
    
    <input type="hidden" name="Meal" value="<%=(Recordset1.Fields.Item("Meal").Value)%>" size="32" />
    
    <input type="hidden" name="Bandana" value="<%=(Recordset1.Fields.Item("Bandana").Value)%>" size="32" />
    <input type="hidden" name="Buff" value="<%=(Recordset1.Fields.Item("Buff").Value)%>" size="32" />
    <input type="hidden" name="DustMask" value="<%=(Recordset1.Fields.Item("DustMask").Value)%>" size="32" />
    <input type="hidden" name="Gita" value="<%=(Recordset1.Fields.Item("Gita").Value)%>" size="32" />
    <input type="hidden" name="SunGlasses" value="<%=(Recordset1.Fields.Item("SunGlasses").Value)%>" size="32" />
    <input type="hidden" name="FreeHugs" value="<%=(Recordset1.Fields.Item("FreeHugs").Value)%>" size="32" />
    
    <input type="hidden" name="OM" value="<%=(Recordset1.Fields.Item("OM").Value)%>" size="32" />
    <input type="hidden" name="Entries" value="<%=(Recordset1.Fields.Item("Entries").Value)%>" size="32" />
    <input type="hidden" name="SuperPack" value="<%=(Recordset1.Fields.Item("SuperPack").Value)%>" size="32" />
    <input type="hidden" name="Checked" value="1" size="32" />
    <br />
    <Br />
    <Br />
    <Br />
    </span>
  <span class="Heading">
  <% If Checked="1" Then 
	  	response.Write("<a href=index_sales.asp>New Search</a>")
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
