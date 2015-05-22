<HTML>
<HEAD>
<TITLE> DesireforAttire ASP Web Application </TITLE>
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
</head>
<body>
<% dim cn, rs, strSQL
if session("username") = "" then
	strSQL = "SELECT * FROM Login WHERE username = '" & request.form("txtUser") & "'"
	'response.write "<BR>" & strSQL
	set cn=server.createobject("ADODB.Connection")
	cn.open "provider=microsoft.jet.oledb.4.0; data source = d:\mageeco\1415\fd1a\adam\desireforattire\products.mdb;"
	set rs=cn.execute(strSQL)
	if rs.EOF and rs.BOF then
		' remember - comparisons are case sensitive!!
		' if no username matches found on database, set session variable to null string
		session("username") = ""
	elseif rs("Password") = request.form("txtPassword") then
			'if password matches, set session variable to username
			session("username") = rs("username")
	else
		'otherwise set to null string
		session("username") = ""
	end if
end if
'if session variable is null string then either username doesn't exist in database or password doesn't match,
'so make user enter details again
if session("username") = "" then
	%>
	<form method= post action="index.asp">
	<table border=2>
	<tr>
		<td>Username:</td><td><input type="text" name="txtUser" size="20">
	</tr>
	<tr>
		<td>Password:</td><td>
        <INPUT TYPE="password" NAME="txtPassword" size="20">
	</tr>
	<tr>
		<td><INPUT TYPE="Submit"></td><td><INPUT TYPE="Reset" Name="">
	</tr>
	</table>
	</form>
<%else%>
		<p>
	<center>
	<table BGColor=#CCFFFF Border=3 Bordercolor=black>
	<tr><td><center><A class="Links" HREF= "listproducts.asp"> List Products </A></CENTER></TD></TR>
	<TR><TD><CENTER><A class="Links" HREF= "addproducts.html"> Add a Product </A></CENTER></TD></TR>
	<TR><TD><CENTER><A class="Links" HREF= "changeproducts.asp"> Update a Product's details </A></CENTER></TD></TR>
	<TR><TD><CENTER><A class="Links" HREF= "delprodcuts.asp"> Delete a Product </A></CENTER></TD></TR>
	</TABLE>
	</CENTER>
<%end if%>
</BODY>
</HTML>