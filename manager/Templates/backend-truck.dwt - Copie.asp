<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- TemplateBeginEditable name="doctitle" -->
<title>Manager</title>
<!-- TemplateEndEditable --><!-- TemplateBeginEditable name="head" --><!-- TemplateEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- TemplateParam name="id" type="text" value="" -->
</head>
<body id="@@(id)@@">
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
<!-- TemplateBeginEditable name="content" -->
 CONTENT 
<!-- TemplateEndEditable --></div></div>
</body>
</html>
