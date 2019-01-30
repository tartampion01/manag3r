<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!-- #include virtual="/includes/check.asp" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

OrderBy = Request.QueryString()

Select Case OrderBy

	Case "unite"
		Set rs=Conn.Execute("Select * from trucks ORDER by unite,marque,intAnnee DESC;")
		par = "par unit&eacute;"
	Case "annee"
		Set rs=Conn.Execute("Select * from trucks ORDER by intAnnee DESC,marque;")
		par = "par ann&eacute;e"
	Case "modele"
		Set rs=Conn.Execute("Select * from trucks ORDER by modele,marque,intAnnee DESC;")
		par = "par mod&egrave;le"
	Case Else
		Set rs=Conn.Execute("Select * from trucks ORDER by succursale,marque,intAnnee DESC;")
		par = "par succursale"
End Select


%>
<html>
<head>
<title>Liste d'inventaire</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
th {
	font-family: "Arial Narrow";
	font-size: 11px;
	color: #FFFFFF;
	background-color: #000000;
}
td {
	font-family: "Arial Narrow";
	font-size: 10px;
	border-top: 1px none #EEEEEE;
	border-right: 1px none #EEEEEE;
	border-bottom: 1px solid #EEEEEE;
	border-left: 1px solid #EEEEEE;
}
.link {
	font-family: "Arial Narrow";
	font-size: 11px;
	color: #FFFFCC;
	text-decoration: none;
}
.link:hover {
	font-family: "Arial Narrow";
	font-size: 11px;
	color: #FF0000;
	text-decoration: underline;
}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td><font size="4" face="Verdana, Arial, Helvetica, sans-serif">Liste d'inventaire&nbsp;<%=par%></font></td>
    <td align="right"><font size="4" face="Verdana, Arial, Helvetica, sans-serif"><%=MakeLocalDate(Date())%></font></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <th nowrap><a title="Trier par unit&eacute;" href="/manager/trucks/liste.asp?unite" class="link">Unit&eacute;</a></th>
    <th nowrap><a title="Trier par suucursale;" href="/manager/trucks/liste.asp?succursale" class="link">Succ.</a></th>
    <th nowrap>Marque</th>
    <th nowrap><a title="Trier par mod&egrave;le" href="/manager/trucks/liste.asp?modele" class="link">Mod&egrave;le</a></th>
    <th nowrap><a title="Trier par ann&eacute;e" href="/manager/trucks/liste.asp?annee" class="link">Ann&eacute;e</a></th>
    <th nowrap># Mois </th>
    <th nowrap>Essieu</th>
    <th nowrap>Vitesse</th>
    <th nowrap>Moteur</th>
    <th nowrap>Pneu</th>
    <th nowrap>Anc. Client</th>
    <th nowrap># S&eacute;rie </th>
    <th nowrap>KM Odo</th>
    <th nowrap>Prix</th>
    <th nowrap>&Eacute;quip.</th>
    <th nowrap>Promo</th>
    <th nowrap>Boni</th>
  </tr>
<%
ct=0
Do While not rs.EOF
If ct mod 2 then
	bgc="bgColor='#EEEEEE'"
Else
	bgc=""
End if
ct=ct+1
%>
  <tr valign="top" <%=bgc%>>
    <td><%=rs("unite")%>&nbsp;</td>
    <td nowrap><%=rs("succursale")%>&nbsp;</td>
    <td nowrap><%=rs("marque")%>&nbsp;</td>
    <td nowrap><%=rs("modele")%>&nbsp;</td>
    <td nowrap><%=rs("intAnnee")%>&nbsp;</td>
    <td align="center" nowrap><%=DateDiff("m",rs("dateAchat"),Date())%></td>
    <td nowrap><%=rs("essieux")%>&nbsp;</td>
    <td><%=rs("transmission")%>&nbsp;</td>
    <td><%=rs("moteur")%>&nbsp;</td>
    <td nowrap><%=rs("pneu_av_dim")%>&nbsp;</td>
    <td><%=rs("ancien_client")%>&nbsp;</td>
    <td nowrap><%=rs("noSerie")%>&nbsp;</td>
    <td nowrap><%=rs("intMillage")%></td>
    <td nowrap><%=rs("prix")%>&nbsp;</td>
    <td><%=rs("equipement2")%>&nbsp;</td>
    <td nowrap><%=rs("promo")%>&nbsp;</td>
    <td nowrap><%=rs("bonis")%>&nbsp;</td>
  </tr>
<%
	rs.moveNext
Loop
%>
</table>

</body>
</html>
