<%Response.Expires=-1441%>
<%
'Admin Only
PageLevel=1
%>
<!--#include virtual="/includes/check.asp" -->
<!-- #include virtual="/intranet/includes/functions.asp" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

ID = Request.QueryString("ID")
IF ID="" then ID=Request.Form("ID")

' New User
If Request.Form<> "" And ID="" then

'	Conn.Execute("LOCK TABLES users WRITE;")
'	Conn.Execute("INSERT users (username) VALUES ('username');")
'	Set rs = Conn.Execute("SELECT ID FROM users WHERE ID IS NULL;")
'	ID = rs("ID")
'	Conn.Execute("UNLOCK TABLES;")
'	Set rs=Nothing

	Conn.Execute("INSERT users (ID) VALUES (NULL);")
	Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;")
	ID = rs("ID")
	Set rs=Nothing	

End IF

' Delete User
IF Request.QueryString("action")="del" then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")
	
	Conn.Execute("Delete from Users Where ID=" & ID & ";")
		
	Conn.Close
	Set Conn=Nothing
	
	Response.Redirect("default.asp")
End IF

' Modify User
If Request.Form("modUser")<>"" then

	strSQL = "Update Users SET" & _
	" FirstName='" & Replace(Request.Form("FirstName"),"'","''") & "'," &_
	" LastName='" & Replace(Request.Form("LastName"),"'","''") & "'," &_
	" username='" & Lcase(Replace(Request.Form("username"),"'","''")) & "'," &_
	" phone='" & Replace(Request.Form("phone"),"'","''") & "'," &_
	" Level=" & Request.Form("level") & "," &_
	" email='" & Replace(Request.Form("email"),"'","''") & "'"

	If Request.Form("password")<> "" then
		If 	Request.Form("password")=Request.Form("password1") then
			strSQL = strSQL & ", password='" & encode(Replace(Request.Form("password"),"'","''")) & "'"
			errMsg = errMsg & "<span class=ok>Password successfully update</span>"
		Else
			errMsg = errMsg & "<span class=err>Password verification failed</span>"
		End If
	End If

	strSQL = strSQL & " WHERE ID=" & ID & ";"

	Conn.Execute(strSQL)
End If

If ID <> "" then
	strSQL = "Select * from users WHERE ID=" & ID & ";"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,Conn
		
	For each item in rs.Fields
		Execute(item.Name & "=rs(""" & item.Name & """)")	
	Next
	rs.Close
End If

Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend-truck.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Gestionnaire Camions Usagés | Réseau Dynamique</title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->

<SCRIPT LANGUAGE="JavaScript">
<!--
function Delete(ID) {
	if (confirm("ATTENTION! Cette opération supprimera cet usager de façon definitive.")){
		string="details.asp?action=del&ID="+ID;
		document.location.href=string;
	}
}
// -->
</SCRIPT>

<!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="" -->
</head>
<body id="">
<div id="container">
<div id="header">
  <ul id="nav">
    <li id="invent_tab"><a href="/manager/trucks/">Camions</a></li>
	<%IF Session("Level")=1 then%>
	<li id="users_tab"><a href="/manager/users/">Utilisateurs</a></li>
	<%End If%>
  </ul>
<ul id="second_nav"><li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li><li><a href="/manager/login/password.asp">Mot de passe</a></li>
<li><a href="/manager/login/logout.asp">Quitter</a></li>
</ul>
</div><div id="content">
<!-- InstanceBeginEditable name="content" -->
 
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr> 
    <td valign="top" class="title"> <h1>Usager: <%=username%> </h1></td>
  </tr>
  <%If errMsg <> "" then%>
  <tr> 
    <td valign="top"><%=ErrMsg%></td>
  </tr>
  <%End IF%>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="3" >
  <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="2"><a href="default.asp"><img src="/manager/images/fup.gif" width="16" height="16" border="0" /></a></th>
  </tr>
  <tr class="titlebar">
    <td align="left" nowrap="nowrap" colspan="2">Identification</td>
  </tr>
  <form method="post" action="/manager/users/details.asp">
    <input type="hidden" name="ID" value="<%=ID%>" />
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Nom:</td>
      <td width="100%" class="cell_content"><input type="text" name="LastName" size="40" value="<%=LastName%>" class="field" /></td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Pr&eacute;nom:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="FirstName" size="40" value="<%=FirstName%>" class="field" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Type:</td>
      <td width="100%" class="cell_content"> 
        <select name="level" class="field">
			
          <option <%If Level=0 then response.write "selected"%> value="0">Usager</option>
			
          <option <%If Level=1 then response.write "selected"%> value="1">Administrateur</option>
        </select>      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Nom de l'utilisateur:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="username" size="40" value="<%=username%>" class="field" />      </td>
    </tr>
    <tr> 
      <td valign="top" align="left" nowrap="nowrap" class="cell_label">Mot de passe:</td>
      <td width="100%" class="cell_content"> 
        <input type="password" name="password" size="15" class="field" />
        <br />
        <input type="password" name="password1" size="15" class="field" />
      (verification)</td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">T&eacute;l&eacute;phone:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="Phone" size="40" value="<%=Phone%>" class="field" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Courriel:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="Email" size="40" value="<%=Email%>" class="field" />
      <a href="mailto:<%=Email1%>"><img src="/manager/images/email.gif" width="23" height="22" align="absmiddle" border="0" alt="Send an email" /></a>	  </td>
    </tr>
    <tr class="titlebar">
      <td align="left" nowrap="nowrap" colspan="2">Actions</td>
    </tr>
    <tr> 
      <td valign="top" align="left" nowrap="nowrap" class="cell_label">&nbsp;</td>
      <td valign="top" width="100%" align="left" class="cell_content"> 
        <input type="submit" value="Enregistrer la fiche" name="modUser" class="button" />
        <% If ID <> "" then%>
        <input onclick="Delete(<%=ID%>)" type="button" value="Effacer l'usager" name="del" class="button" />
        <%End If%>      </td>
    </tr>
  </form>
</table> 
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>