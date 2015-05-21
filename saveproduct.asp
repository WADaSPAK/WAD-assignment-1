<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Saved Changed Product Details</TITLE>
<META NAME="Generator" CONTENT="">
<style type="text/css">
body { font-family: verdana;
	   background-color: #1696f7;
	   font-size: 10px;
	   colour: #000000;
}

TD	{
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

a.Links	{color: yellow;
		 font-weight: 0;

}


</style>
</HEAD>

<BODY>
<% dim cn, strSQL
Response.write "The record for " & request.form("Name")
Response.write " " & request.form("Name")
Response.write " has been updated on our Database."
StrSQL = "UPDATE PRODUCTS SET NAME='" & request.form("Name") & "'," _
	& " Type = '" & request.form("Type") & "'," _
	& " Size = '" & request.form("Size") & "'," _
	& " Colour = '" & request.form("Colour") & "'," _
	& " Stock = '" & request.form("Stock") & "'," _	
	& " Price = '" & request.form("Price") & "'," _
	& "'," _
	& "' WHERE NAME=" _
	& request.form("Name")
response.write "<BR>" & strSQL
set cn = server.createobject("ADODB.Connection")
cn.open "provider=microsoft.jet.oledb.4.0; data source = products.mdb;"
cn.execute strsql
cn.close
set cn = nothing
%>
<p>
<A HREF="productlist.asp">Product List</A>
</BODY>
</HTML>