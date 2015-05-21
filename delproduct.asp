<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Delete a Product</TITLE>

</HEAD>

<BODY bgcolor = #66CCFF>
<CENTER>
<H2> Select the Product you wish to delete (by Name) <%=Date%></H2>
<TABLE bgcolor = #FF0000 border = 2>
<TR>
	<TH>Name</TH><TH>Type</TH><TH>Size</TH><TH>Colour</TH><TH>Stock</TH><TH>Price</TH>
</TR>
<% dim cn, rs
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source=products.mdb;"
set rs=cn.execute("Select * From Products")
DO until rs.EOF
	response.write "<TR>"
	response.write "<TD><a href='removeproduct.asp?name=" & rs("name") & "'>Name=" & rs("name") & "</TD>"
	response.write "<TD>" & rs("Name") & "</TD>"
	response.write "<TD>" & rs("Type") & "</TD>"
	response.write "<TD>" & rs("Size") & "</TD>"
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
</BODY>
</HTML>