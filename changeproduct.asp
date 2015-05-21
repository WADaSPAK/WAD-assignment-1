<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Change Product Details </TITLE>
<style type="text/css">
body { font-family: verdana;
	   background-color: #1696f7;
	   font-size: 10px;
	   colour: #FFFFFF;
}

TH	{
	color: #FFFFFF;
	font-size: 16px;
	font-weight: bold;
}

TABLE	{
	background-color: #003366;
	border: none;
}

a {
	color: yellow;
}

H2	{
	color: #FFCC00;

}


</style>
</HEAD>

<BODY bgcolor = #66CCFF>
<CENTER>
<H2> Select the Product you wish to change (by Name) <%=Date%></H2>
<TABLE bgcolor = #FF0000 border = 2>
<TR>
	<TH>Name</TH><TH>Type</TH><TH>Size</TH><th>Colour</th><th>Stock</th><th>Price</th>
</TR>
<% dim cn, rs
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source = D:\herrons\FD1B\Contacts\contacts.mdb;"
set rs=cn.execute("Select * From Contacts")
DO until rs.EOF
	response.write "<TR>"
	response.write "<TD><a href='updatecontact.asp?id=" & rs("name") & "'>" & rs("Name") & "</TD>"
	response.write "<TD>" & rs("Type") & "</TD>"
	response.write "<TD>" & rs("Size") & "</a></TD>"
	response.write "<TD>" & rs("Colour") & "</TD>"
	response.write "<TD>" & rs("Stock") & "</TD>"
	response.write "<TD>" & rs("Price") & "</TD>"
	response.write "</TR>"
	rs.movenext
loop
rs.close
set rs = nothing
cn.close
set cn = nothing
%>
</TABLE>
</CENTER>
<p>
<A HREF ="admin.html"> Home Page </A>
</BODY>
</HTML>