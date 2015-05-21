<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Delete Product? </TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
function goBack()
{
	document.location.href = "delproduct.asp"
}
function deleters()
{
	document.frmDelete.action = "confirmdelete.asp"
	document.frmDelete.submit()
}
//-->
</SCRIPT>
</HEAD>
<BODY BGColor = #FF99FF>
<CENTER>
<H2>Delete this Product?</H2>

<%dim cn, rs, strSQL
strSQL = "SELECT * FROM PRODUCTS WHERE Name = " & request.querystring("name")
set cn=server.createobject ("adodb.connection")
cn.open"provider=microsoft.jet.oledb.4.0; data source=products.mdb;"
set rs = cn.execute(strSQL)%>
<FORM METHOD=POST ACTION="confirmdelete.asp" NAME="frmDelete">
<TABLE bgcolor = #CCCC00 border = 2>
<TR>
	<TD>Name</TD><TD><%=rs("Name")%></TD>
</TR>
<TR>
	<TD>Type</TD><TD><%=rs("Type")%></TD>
</TR>
<TR>
	<TD>Size</TD><TD><%=rs("Size")%></TD>
</TR>
<TR>
	<TD>Colour</TD><TD><%=rs("Colour")%></TD>
</TR>
<TR>
	<TD>Stock</TD><TD><%=rs("Stock")%></TD>
</TR>
<TR>
	<TD>Price</TD><TD><%=rs("Price")%></TD>
</TR>
<TR>
	<TD><Input Type="submit" value="Delete Product">
	<input type="hidden" name="Name" value="<%=rs("name")%>"></TD>
	<TD><Input Type="button" value=" Don't Delete Product" onClick="goBack()"></TD>
</TR>
</TABLE>
</FORM>
<%rs.close
set rs = nothing
cn.close
set cn = nothing%>
</CENTER>
<A HREF ="admin.html"> Home Page </A>
</BODY>
</HTML>
