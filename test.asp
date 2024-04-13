<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/HoliCon.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_HoliCon_STRING
Recordset1_cmd.CommandText = "SELECT * FROM SpanishFork2015" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Spanish Fork Festival of Colors</title>
<link href="styler.css" rel="stylesheet" type="text/css">
</head>

<body>
<p>Test for the ASP Running <br />
  <br />
<span class="Heading">
  <%
=(Recordset1.Fields.Item("LastName").Value)%>, 
    
  <%
=(Recordset1.Fields.Item("FirstName").Value)%>
</span></p>
<span class="Text">
<% Dim Checked
 Checked=(Recordset1.Fields.Item("Checked").Value)
	If Checked="1" Then response.Write("<style=CSS 		class=text_ckd>")
	%>
<% Colors =(Recordset1.Fields.Item("Colors").Value)
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
%>
<% If Colors<>0  Then Response.Write("Color Bags: " & Colors &"</br>")%>
<% If TShirtH<>0  Then Response.Write("Holi T-Shirt: " & TShirtH &"</br>")%>
<% If TShirtG<>0  Then Response.Write("Ganesh T-Shirt: " & TShirtG &"</br>")%>
<% If TShirtK<>0  Then Response.Write("Krishna T-Shirt: " & TShirtK &"</br>")%>
<% If TShirtOM<>0  Then Response.Write("T-Shirt OM: " & TShirtOM &"</br>")%>
<% If TShirtOmDlx<>0  Then Response.Write("T-Shirt OM-DLX: " & TShirtOmDlx &"</br>")%>
<% If Bandana<>0  Then Response.Write("Bandana: " & Bandana &"</br>")%>
<% If Buff<>0  Then Response.Write("Buff: " & Buff &"</br>")%>
<% If Gita<>0  Then Response.Write("Bhagwad Gita: " & Gita &"</br>")%>
<% If Mask<>0  Then Response.Write("Dust Masks: " & Mask &"</br>")%>
<% If SunGlasses<>0  Then Response.Write("Sun Glasses: " & SunGlasses &"</br>")%>
<% If Meal<>0  Then Response.Write("Hot Meal: " & Meal &"</br>")%>
<% If Entries="Multiple"  Then Response.Write("Entries: " & Entries &"</br>")%>

</span>
</span>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
