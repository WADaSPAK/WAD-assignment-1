<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Product List </TITLE>

</HEAD>

<BODY bgcolor = #66CCFF>
<CENTER>
<H2>Product List <%=Date%></H2>
<TABLE bgcolor = #FF0000 border = 2>
<TR>
	<TH>Name</TH><TH>Type</TH><TH>Size</TH><th>Colour</th>
	<TH>Stock</TH><th>Price</th>
</TR>
<% dim cn, rs
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source=d:\mageeco\1415\fd1a\adam\DesireForAttire\products.mdb;"
set rs=cn.execute("Select * From Products")
DO until rs.EOF
	response.write "<TR>"
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
<A HREF ="addproduct.htm"> Add another Product</A>
<A HREF ="delproduct.asp"> Delete a Product</A>
<A HREF ="changeproduct.asp"> Update a Product's Details</A>
</BODY>
</HTML>