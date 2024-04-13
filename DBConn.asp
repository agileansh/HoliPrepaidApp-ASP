<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%

Dim MM_Holi

dsn_name="\HoliDB.accdb"               'name of database file
sdsndir=Server.MapPath("dbf")	   'path of database file

db_path=sdsndir & dsn_name

MM_Holi =  "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & db_path
%>
