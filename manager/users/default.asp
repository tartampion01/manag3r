<%Response.Expires=-1441%>
<%
'Admin Only
PageLevel=1
%>
<!--#include virtual="/includes/check.asp" -->
<%
tableName = "users"
Title = "Liste des utilisateurs"
SearchItems=""

dim Headers(3,5)
Headers(0,0)="LastName"
Headers(0,1)="Nom"
Headers(0,2)=false
Headers(0,3)="rs(""FirstName"") & ""&nbsp;"" & rs(""LastName"")"

Headers(1,0)="phone"
Headers(1,1)="Téléphone"
Headers(1,2)=false

Headers(2,0)="email"
Headers(2,1)="Courriel"
Headers(2,2)=false

Headers(3,0)="Level"
Headers(3,1)="Type d'utilisateur"
Headers(3,2)=false
Headers(3,3)="UserLevel(rs(""Level""))"

strSQL = "Select * from users "

overdue_condition = "false"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend-truck.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Gestionnaire Camions Usagés | Réseau Dynamique</title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="users" -->
</head>
<body id="users">
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
 
<!--#include virtual="/intranet/includes/list_inc.asp" -->
 
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>

