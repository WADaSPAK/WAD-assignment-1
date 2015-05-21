<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Desire For Attire</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<% dim cn, rs, strSQL
if session("Email") = "" then
	strSQL = "SELECT * FROM login WHERE Email = '" & request.form("txtEmail") & "'"
	'response.write "<BR>" & strSQL
	set cn=server.createobject("ADODB.Connection")
	cn.open "provider=microsoft.jet.oledb.4.0; data source = products.mdb;"
	set rs=cn.execute(strSQL)
	if rs.EOF and rs.BOF then
		' remember - comparisons are case sensitive!!
		' if no email matches found on database, set session variable to null string
		session("Email") = ""
	elseif rs("Password") = request.form("txtPassword") then
			'if password matches, set session variable to email
			session("Email") = rs("Email)
	else
		'otherwise set to null string
		session("Email") = ""
	end if
end if
'if session variable is null string then either email doesn't exist in database or password doesn't match,
'so make user enter details again
if session("Email") = "" then
	%>

<%end if%>
</body>
