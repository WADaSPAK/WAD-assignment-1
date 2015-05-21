<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Record Deleted</TITLE>
</HEAD>
<BODY bgColor=#FFCCCC>
<% dim cn, strSQL
Response.write "The record with Name = " & request.form("Name")
Response.write " has been removed from the Products Database"
StrSQL = "DELETE FROM Products WHERE Name=" & request.form("Name")
response.write "<BR>" & strSQL
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source=products.mdb;"
cn.execute strsql
cn.close
set cn = nothing
%>
<p>
<A HREF="productlist.asp">Products List</A>
</BODY>
</HTML>
