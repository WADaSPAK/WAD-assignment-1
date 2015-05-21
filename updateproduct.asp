<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Update a Product </TITLE>
<style type="text/css">
body { font-family: verdana;
	   background-color: #1696f7;
	   font-size: 10px;
	   colour: #FFFFFF;
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


</style>

</HEAD>
<BODY>
<CENTER>
<H2>Update a Products details</H2>

<%dim cn, rs, strSQL
strSQL = "SELECT * FROM PRODUCTS WHERE NAME = " & request.querystring("Name")
set cn=server.createobject ("adodb.connection")
cn.open"provider=microsoft.jet.oledb.4.0; data source = products.mdb;"
set rs = cn.execute(strSQL)%>
<FORM METHOD=POST ACTION="saveproduct.asp" NAME="frmupdate">
<TABLE>
<TR>
	<TD>Name</TD><TD>
    <Input Type = "text" Name = "Name" value="<%=rs("Name")%>" size="20"></TD>
</TR>
<TR>
	<TD>Type</TD><TD>
    <Input Type = "text" Name = "Type" value="<%=rs("Type")%>" size="20"></TD>
</TR>
<TR>
	<TD>Size</TD><TD>
    <Input Type = "text" Name = "Size" value="<%=rs("Size")%>" size="20"></TD>
</TR><TR>
	<TD>Colour</TD><TD>
    <Input Type = "text" Name = "Colour" value="<%=rs("Colour")%>" size="20"></TD>
</TR>
<TR>
	<TD>Stock</TD><TD>
    <Input Type = "text" Name = "Stock" value="<%=rs("Stock")%>" size="20"></TD>
</TR>
	<TD>Price</TD><TD>
    <Input Type = "text" Name = "Price" value="<%=rs("Price")%>" size="20"></TD>
</TR>
<TR>
	<TD><Input Type="submit" value="Update Product">&nbsp;
	<TD><input type="hidden" name="txtid" value="<%=rs("id")%>">
	</TD>
</TR>
</TABLE>
</FORM>
<%rs.close
set rs = nothing
cn.close
set cn = nothing%>
</CENTER>
</BODY>
</HTML>