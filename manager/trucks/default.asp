<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!--#include virtual="/includes/check.asp" -->
<%

'Table to display
tableName = "trucks"
Title = "Les Camions"
SearchItems="modele,marque,intAnnee,unite,succursale"

'Fields to display
dim Headers(5,5)

Headers(0,0)="unite"
Headers(0,1)="Unité No"
Headers(0,2)=false

Headers(1,0)="succursale"
Headers(1,1)="Succ."
Headers(1,2)=false

Headers(2,0)="marque"
Headers(2,1)="marque"
Headers(2,2)=false
'Headers(1,3)="rs(""Category"")"
'Headers(1,4)="Select ID as Category_ID,Category from Category WHERE Level=1;"
'Headers(1,5)="rsCombo(""Category"")"

Headers(3,0)="modele"
Headers(3,1)="Modèle"
Headers(3,2)=false

Headers(4,0)="intAnnee"
Headers(4,1)="Année"
Headers(4,2)=false

Headers(5,0)="intStatus"
Headers(5,1)="Fiche"
Headers(5,2)=true
dim Status(2)
	status(0)="Inactive"
	status(1)="Active"
Headers(5,3)="status(rs(""intStatus""))"
Headers(5,4)="Select distinct(intStatus) from trucks;"
Headers(5,5)="status(rsCombo(""intStatus""))"

Headers(6,0)="daysin"
Headers(6,1)="# jours"
Headers(6,2)=false
Headers(6,3)="ShowDays(rs(""daysin""))"

overdue_condition = "rs(""resUsers_ID"")"

'strSQL = "SELECT * FROM trucks "
strSQL = "SELECT *,cast(ifnull(DATEDIFF(CURDATE(),dateAchat),0) as SIGNED) as daysin from inventory FROM trucks "


print_action = "href=""liste.asp"" target=""_blank"""

function ShowDays(d)
	ShowDays = "N/D"
	if cint(d)>0 then
		ShowDays = "<span class='err'>" & d & "</span>"
	end if
end function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend-truck.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Gestionnaire Camions Usagés | Réseau Dynamique</title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="inventory" -->
</head>
<body id="inventory">
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
