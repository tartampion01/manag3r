<!-- #include virtual="/includes/check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend-truck.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Gestionnaire Camions Usagés | Réseau Dynamique</title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/includes/styles_app.css" rel="stylesheet" type="text/css" />
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
<table width="100%%" border="0" cellspacing="0" cellpadding="0" height="80%">
  <tr>
    <td width="100%" height="100%" align="center" valign="top">
            <p><font face="Tahoma"><span style="font-size:17px;"><br />Bienvenue <%=Session("FullName")%>!</span></font></p>
            <p><img src="images/header_2008.gif" alt="Le R&eacute;seau Dynamque" width="841" height="182" /></p>
    <p>&nbsp;</p></td>
  </tr>
</table>
<h3>&nbsp;</h3>
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
