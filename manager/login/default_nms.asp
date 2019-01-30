<!-- #include virtual="/manager/includes/functions.asp" -->
<%


'if request.QueryString("k")<>"" then
'
'	Set Conn = Server.CreateObject("ADODB.Connection")
'	Conn.Open Application("cstring")
'
'	Set rs=Conn.Execute("SELECT id,password,username from users;")
'	
'	Do while not rs.EOF
'		response.write rs("username")
'		Conn.Execute("update users set password='"& encode(rs("password")) &"' WHERE id='"& rs("ID") &"';")
'		rs.MoveNext
'	loop
'	
'end if

UserName = Request.Form("username")
Password = Request.Form("password")
URL = Request.QueryString("url")

IF URL="" then URL = "/manager/"

If UserName <>"" then

	'Revoke Admin rights
	Session("Level")=""
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")

	Set rs=Conn.Execute("SELECT * from users WHERE username='" & UserName & "';")

	If NOT rs.EOF then
		Session("managerID")= rs("ID")
		Session("username")= rs("username")
		Session("fullname")= rs("FirstName") & " " & rs("LastName")
		If Password = rs("password") then
			If Request.Form("cookie")="ON" then
				Response.Cookies("managerID") = Session("managerID")
				Response.Cookies("managerID").Expires = Date + 30
				Response.Cookies("username") = Session("username")
				Response.Cookies("username").Expires = Date + 30
				Response.Cookies("fullname") = Session("fullname")
				Response.Cookies("fullname").Expires = Date + 30
			Else
				Response.Cookies("managerID").Expires = Date - 1000
				Response.Cookies("username").Expires = Date - 1000
				Response.Cookies("fullname").Expires = Date - 1000
			End If
			Response.Redirect(URL)
		Else
			Response.Redirect("/manager/login/?err=1&URL=" & Server.URLEncode(URL))
		End If
	Else
		Response.Redirect("/manager/login/?err=2&URL=" & Server.URLEncode(URL))
	End If	

	Set rs=Nothing
	Conn.Close
	Set Conn=Nothing
End If

If Request.Cookies("managerID")<>"" then
	cookieCheck = "checked"
End If

If Request.QueryString("err")=3 then
	ErrMsg = "Vous ne pouvez accéder à cette section."
End If
If Request.QueryString("err")=2 then
	ErrMsg = "Nom d'utilisateur invalide"
End If
If Request.QueryString("err")=1 then
	ErrMsg = "Mot de passe invalide"
End If


%>
<html>
 
<head>
 
<title>Gestionnaire Camions Usagés | Réseau Dynamique | Login</title>
<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW">
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#FFFFFF">
 
<table border="0" cellspacing="0" width="100%" cellpadding="0" height="100%">
  <tr> 
    <td valign="top" align="center"> 
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <td align="center" valign="middle">
                            <img src="../images/header_2008.gif" alt="Le réseau Dynamque" width="841" height="182">
                      </td>
                    </tr>
                </table>
      <form method="POST" action="default_nms.asp?URL=<%=Server.URLEncode(URL)%>">
        <table width="250" border="0" cellpadding="0" cellspacing="5">
          <tr>
            <td align="left"><h2 align="center">Bienvenue!</h2>
            <span class="err"><%=ErrMsg%></span>
            </td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0" id="login">
                <tr>
                  <th><strong>Acc&egrave;s utilisateur </strong></th>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="left"><strong>Utilisateur:</strong></td>
                        <td align="right"><input type="text" name="username" value="<%=session("username")%>" style="width:150px" />                        </td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap"><strong>Mot de passe :</strong></td>
                        <td align="right"><input type="password" name="password" style="width:150px" /></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td><table border="0" width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><input type="checkbox" name="cookie" value="ON" />                        </td>
                        <td align="left">M&eacute;moriser </td>
                        <td align="right"><input name="submit" type="submit" class="button" value="Login" />                        </td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><a href="retreive.asp">Mot de passe oubli&eacute;?</a></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table> 
</body>

</html>
