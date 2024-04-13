<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/HoliCon.asp" -->
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
    MM_editCmd.CommandText = "UPDATE SpanishFork2015 SET ID = ?, LastName = ?, FirstName = ?, GuestLast = ?, GuestFirst = ?, Colors = ?, TshirtsHoli = ?, HoddieK = ?, Ganesh = ?, OmDlx = ?, Meal = ?, Bandana = ?, Buff = ?, DustMask = ?, Gita = ?, SunGlasses = ?, FreeHugs = ?, OM = ?, Entries = ?, SuperPack = ?, Checked = ? WHERE ID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("ID"), Request.Form("ID"), null)) ' adDouble
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
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_HoliCon_STRING
Recordset1_cmd.CommandText = "SELECT * FROM SpanishFork2015 WHERE ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<% Lname=(request.QueryString("LastName"))
FName=(Request.QueryString("FirstName")) 
response.Write(Lname & " ")
Response.Write(FName)
Response.Write(request.QueryString("Unique"))
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
<meta name="viewport" content="width=device-width; initial-scale=1; minimal-ui">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width" />
<title>Untitled Document</title>
</head>

<body>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">ID:</td>
      <td><input type="text" name="ID" value="<%=(Recordset1.Fields.Item("ID").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">LastName:</td>
      <td><input type="text" name="LastName" value="<%=(Recordset1.Fields.Item("LastName").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">FirstName:</td>
      <td><input type="text" name="FirstName" value="<%=(Recordset1.Fields.Item("FirstName").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">GuestLast:</td>
      <td><input type="text" name="GuestLast" value="<%=(Recordset1.Fields.Item("GuestLast").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">GuestFirst:</td>
      <td><input type="text" name="GuestFirst" value="<%=(Recordset1.Fields.Item("GuestFirst").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Colors:</td>
      <td><input type="text" name="Colors" value="<%=(Recordset1.Fields.Item("Colors").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">TshirtsHoli:</td>
      <td><input type="text" name="TshirtsHoli" value="<%=(Recordset1.Fields.Item("TshirtsHoli").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">HoddieK:</td>
      <td><input type="text" name="HoddieK" value="<%=(Recordset1.Fields.Item("HoddieK").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Ganesh:</td>
      <td><input type="text" name="Ganesh" value="<%=(Recordset1.Fields.Item("Ganesh").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">OmDlx:</td>
      <td><input type="text" name="OmDlx" value="<%=(Recordset1.Fields.Item("OmDlx").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Meal:</td>
      <td><input type="text" name="Meal" value="<%=(Recordset1.Fields.Item("Meal").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Bandana:</td>
      <td><input type="text" name="Bandana" value="<%=(Recordset1.Fields.Item("Bandana").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Buff:</td>
      <td><input type="text" name="Buff" value="<%=(Recordset1.Fields.Item("Buff").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">DustMask:</td>
      <td><input type="text" name="DustMask" value="<%=(Recordset1.Fields.Item("DustMask").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Gita:</td>
      <td><input type="text" name="Gita" value="<%=(Recordset1.Fields.Item("Gita").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SunGlasses:</td>
      <td><input type="text" name="SunGlasses" value="<%=(Recordset1.Fields.Item("SunGlasses").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">FreeHugs:</td>
      <td><input type="text" name="FreeHugs" value="<%=(Recordset1.Fields.Item("FreeHugs").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">OM:</td>
      <td><input type="text" name="OM" value="<%=(Recordset1.Fields.Item("OM").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Entries:</td>
      <td><input type="text" name="Entries" value="<%=(Recordset1.Fields.Item("Entries").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SuperPack:</td>
      <td><input type="text" name="SuperPack" value="<%=(Recordset1.Fields.Item("SuperPack").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Checked:</td>
      <td><input type="text" name="Checked" value="<%=(Recordset1.Fields.Item("Checked").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="Update record" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1" />
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>" />
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
