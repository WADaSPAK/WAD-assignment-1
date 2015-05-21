<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Add a Product</TITLE>
<META NAME="Generator" CONTENT="EditPlus">

</HEAD>

<BODY>
<% dim cn, strSQL
Response.write "The record for " & request.form("Name")
Response.write " " & request.form("Type")
Response.write " have been booked."
strSQL = "Insert Into contacts(Name, Type, Size, Colour, Stock, Price) " & "VALUES('"& request.queryString("Name") & "','" & request.queryString("Type") & "','" & request.queryString("Size") & "','" & request.queryString("Colour") & "','" & request.queryString("Stock") & "','" & request.queryString("Price") & "')"
response.write strSQL
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source=d:\mageeco\1415\fd1a\adam\desireforattire\products.mdb;"
cn.execute strSQL
cn.close
set cn = nothing
%>
<p>
<a href="admin.html">Click here to return to main menu.</a>
</BODY>
</HTML>
