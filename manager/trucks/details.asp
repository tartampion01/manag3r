<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!--#include virtual="/includes/check.asp" -->
<%

Function ConvertFileToBase64( file )

		' This script reads jpg picture converts it to base64
		' code using encoding abilities of MSXml2.DOMDocument object and saves

		Const fsDoOverwrite     = true  ' Overwrite file with base64 code
		Const fsAsASCII         = false ' Create base64 code file as ASCII file
		Const adTypeBinary      = 1     ' Binary file is encoded

		' Variables for writing base64 code to file
		Dim objFSO
		Dim objFileOut

		' Variables for encoding
		Dim objXML
		Dim objDocElem

		' Variable for reading binary picture
		Dim objStream

		' Open data stream from picture
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Type = adTypeBinary
		objStream.Open()

		objStream.LoadFromFile(file)

		' Create XML Document object and root node
		' that will contain the data
		Set objXML = CreateObject("MSXml2.DOMDocument")
		Set objDocElem = objXML.createElement("Base64Data")
		objDocElem.dataType = "bin.base64"

		' Set binary value
		objDocElem.nodeTypedValue = objStream.Read()

		ConvertFileToBase64 = objDocElem.text

		' Clean all
		Set objFSO = Nothing
		Set objFileOut = Nothing
		Set objXML = Nothing
		Set objDocElem = Nothing
		Set objStream = Nothing
	End Function
	
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

Set RequestForm = Server.CreateObject("ABCUpload4.XForm")
RequestForm.Overwrite = True
RequestForm.AbsolutePath = False
RequestForm.MaxUploadSize = 999999999

ID = Request.QueryString("ID")
IF ID="" then ID=RequestForm("ID")


' New Record
If (RequestForm<>"" And ID="") OR RequestForm("_clone")<>"" then

	Conn.Execute("INSERT trucks (ID) VALUES (NULL);")
	Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;")
	ID = rs("ID")
	Set rs=Nothing

End IF

' Delete Doc
IF request.querystring("action")="del" then
	set rs = conn.execute("SELECT name FROM pictures WHERE product_id='"&id&"'")
	do while not rs.EOF
		FSODelete("/pics/lg/"&rs("name"))
		FSODelete("/pics/"&rs("name"))
		FSODelete("/pics/small/"&rs("name"))
		rs.movenext
	loop
	Conn.Execute("Delete from trucks Where ID=" & ID & ";")
	Response.Redirect("default.asp")
End IF

'deletespec
IF request.querystring("action")="delspec" then
	FSODelete("/pics/specs/spec"&request.querystring("id")&".pdf")
	response.Redirect("details.asp?id=" & request.querystring("id"))
End IF

'delete image
if request.querystring("picid") <> "" then
	Conn.execute("Delete from pictures WHERE id="&request.querystring("picid"))
		FSODelete("/pics/lg/pic"&request.querystring("picid")&".jpg")
		FSODelete("/pics/pic"&request.querystring("picid")&".jpg")
		FSODelete("/pics/small/pic"&request.querystring("picid")&".jpg")
	set rs = nothing
	response.Redirect("details.asp?id=" & request.querystring("id"))
end if

'delete doc
if request.querystring("docid") <> "" then
	Conn.execute("Delete from docs WHERE id="&request.querystring("docid"))
		FSODelete("/pics/docs/doc"&request.querystring("docid")&".pdf")
	set rs = nothing
	response.Redirect("details.asp?id=" & request.querystring("id"))
end if

' Modify Doc
If RequestForm<>"" then

	strSQL= "Update trucks SET "
	Separator=""
	 For each item in RequestForm
		If item <> "ID" and Left(item,1)<>"_" then
			If InStr(item,"Date")<>0 then
				If Requestform(item)<>"" then
				 	strSQl = strSQL & Separator & item & "='" & Year(Requestform(item))&"/"&Month(Requestform(item))&"/"&Day(Requestform(item)) & "'"
				End IF
			Else
				If Left(item,3)="int" then
					IF IsNumeric(RequestForm(item)) Then
						strSQl = strSQL & Separator & item & "=" & RequestForm(item)
					Else
						strSQl = strSQL & Separator & item & "=Null"
					End IF
				Else
					strSQl = strSQL & Separator & item & "='" & Replace(Replace(RequestForm(item),"""","&quot;"),"'","''") & "'"
				End If
			End IF
			Separator = ", "
		End If
	 Next
	 
	'Active checkbox
	 IF RequestForm("intStatus")="" then strSQL = strSQL & ",intStatus=0"
	 
	 'IF Session("Level")< 1 then strSQL= "Update trucks SET intStatus=intStatus "
	 
	 ' Reserver
	 If RequestForm("_Reserve") <> "" And IsDate(RequestForm("_Date")) and RequestForm("_ResUsers_ID") <> "0" then
	  	strSQL = strSQL & ",resDateHeure='" & RequestForm("_Date") & " " & RequestForm("_HH") & ":" & RequestForm("_MM") & ":00'"
		strSQL = strSQL & ",resUsers_ID=" & RequestForm("_ResUsers_ID")	
	 End If
	 
	 If RequestForm("_Cancel") <> "" then
	  	strSQL = strSQL & ",resDateHeure=Null"
		strSQL = strSQL & ",resUsers_ID=0"	
	 End If 
	
	 strSQL = strSQL & " Where ID = " & ID & ";"
	'Response.Write Server.HTMLEncode(strSQL)
	'Response.End
	 Conn.Execute(strSQL)
	 
	 If RequestForm("_clone")<>"" then
	 	Conn.Execute("UPDATE trucks SET unite = CONCAT(unite,' COPIE') Where ID = " & ID & ";")
		'Call CopyPictures(RequestForm("ID"),ID,"/product_images/")
	 End IF
	 

	Function ResizeImage(srcFileName,destFileName,MaxWidth,MaxHeight,Crop)
		'Reduce and compress large image
		
		on error resume next
		
		Dim Image
		
		srcFileName = Server.MapPath(srcFileName)
	
		Set Image = Server.CreateObject("GflAx.GflAx")
		
		Image.LoadBitmap srcFileName
		Image.SaveJPEGQuality = 60
		
		'ignore  X or Y modes
		IF MaxWidth=0 then MaxWidth=Image.Width
		IF MaxHeight=0 then MaxHeight=Image.Height 
		
'		if crop then
'			
'			
'							
'			Response.write  (Image.Width - MaxWidth)/2 & "," & (Image.Height - MaxHeight)/2 & "," & MaxWidth & "," & MaxHeight
'			response.end
'			Image.Crop (Image.Width - MaxWidth)/2,(Image.Height - MaxHeight)/2,MaxWidth,MaxHeight
'			
'		else
	
			If Image.Width > MaxWidth then
				Image.Resize MaxWidth,Int(Image.Height*MaxWidth/Image.Width) 
			End IF
			If Image.Height > MaxHeight then
				Image.Resize Int(Image.Width*MaxHeight/Image.Height),MaxHeight
			End If			
			
		'end if

		Image.SaveBitmap Server.MapPath(destFileName)
	
		
		'End if
	
		if err.number <> 0 then
			response.write "An error occured creating Image:" & err.description & "<br>"
			response.end 
		End If
	
		Set Image = Nothing
		
		'response.end 
		
	End Function
 

	'image order
	pic_order = 0
	for each pic_id in RequestForm("_picture_id")
		 Conn.Execute("Update pictures SET intorder='" & pic_order & "' WHERE id='"& pic_id &"';")	
		 pic_order=pic_order + 1
	next	 
	 
	'image upload
	for each items in RequestForm("_img_file")
		Set theField = items
		If theField.FileExists and theField.ImageType <> 0 Then
			Conn.Execute("INSERT pictures (ID,product_id) VALUES (NULL,'"&ID&"');")
			Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;") 
			picID = rs("ID")
			picfilename =  "pic" & picID & ".jpg"
			theField.Save ("/pics/tmp/" & picfilename)
			Call ResizeImage("/pics/tmp/"&picfilename,"/pics/lg/"&picfilename,800,0,false)
			Call ResizeImage("/pics/tmp/"&picfilename,"/pics/"&picfilename,300,0,false)
			Call ResizeImage("/pics/tmp/"&picfilename,"/pics/small/"&picfilename,80,80,true)
	
			'To insert image as base64 into DB
			Dim b64
			b64 = "data:image/jpeg;base64," & ConvertFileToBase64(Server.MapPath("/pics/tmp/"&picfilename))
			
			Conn.Execute("UPDATE pictures SET intorder='" & pic_order & "', name='"&picfilename &"', base64_picture='" & b64 & "' WHERE ID ='"&picID&"';")
		
			pic_order=pic_order + 1		
			FSODelete("/pics/tmp/"&picfilename) 
			
		end if
		set rs = nothing
	next

	'doc upload
	for each items in RequestForm("_doc_file")
		Set theField = items
		If theField.FileExists Then
			Conn.Execute("INSERT docs (ID,product_id) VALUES (NULL,'"&ID&"');")
			Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;") 
			docID = rs("ID")
			docfilename =  "doc" & docID & ".pdf"
			theField.Save ("/pics/docs/" & docfilename)
			Conn.Execute("UPDATE docs SET intorder='" & doc_order & "', name='"&theField.FileName&"' WHERE ID ='"&docID&"';")
		end if
		set rs = nothing
	next

	 
	for each items in RequestForm("_spec")
		Set theField = items
		if theField.FileExists Then
			theField.Save ("/pics/specs/spec" & ID & ".pdf")
		end if
	next
	 
	 
	 
	 Response.Redirect("details.asp?ID="&ID)
End If

'HH MM combos
For i = 0 to 23
	hhCombo = hhCombo & "<option value=" & i & ">" & Right(formatNumber(i/100,2),2) & "</option>" & VbCrLf
Next
For i = 0 to 59
	mmCombo = mmCombo & "<option value=" & i & ">" & Right(formatNumber(i/100,2),2) & "</option>" & VbCrLf
Next

If ID <> "" then

	strSQL = "Select * from trucks WHERE ID=" & ID & ";"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,Conn
		
	For each item in rs.Fields
		Execute(item.Name & "=rs(""" & item.Name & """)")
	Next
	rs.Close
	


Else
	hh = Hour(Now())
	mm = Minute(Now())
	dateAchat = date()
End If


hhCombo = Replace(hhCombo,"="& hh &">","="& hh &" selected>")
mmCombo = Replace(mmCombo,"="& mm &">","="& mm &" selected>")







%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend-truck.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
 
<title>Gestionnaire Camions Usagés | Réseau Dynamique</title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->

<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
<link rel="stylesheet" type="text/css" href="../includes/jquery/lightbox/lightbox.css" />
<script type="text/javascript" src="../includes/jquery/lightbox/lightbox.min.js"></script>
<style>
 .lb-nav {
   display: none !important;
 }
</style>
<SCRIPT language="JavaScript">
<!--
function Delete(ID) {
	if (confirm("ATTENTION! Ceci effacera cette entrée définitivement.")){
		string="details.asp?action=del&ID="+ID;
		document.location.href=string;
	}
}

function DeletePic(id,picid) {
	if (confirm("Voulez-vous supprimer la photo?")){
		window.location = "details.asp?id="+id+"&picid="+picid;
	}
}

function DeleteDoc(id,docid) {
	if (confirm("Voulez-vous supprimer le document?")){
		window.location = "details.asp?id="+id+"&docid="+docid;
	}
}

function DeleteSpec(id) {
	if (confirm("Voulez-vous supprimer la fiche descriptive?")){
		string="details.asp?action=delspec&ID="+id;
		document.location.href=string;
	}
}

$(function(){
	$("#addpic").click(function(){
		$("#browsetd").prepend('<input type="file" name="_img_file"/><br>');
	});
	
	$("#adddoc").click(function(){
		$("#browsedoctd").prepend('<input type="file" name="_doc_file"/><br>');
	});
	
	$("#allpics").sortable();
});

// -->
</SCRIPT>

<link rel="stylesheet" type="text/css" href="/includes/spiffyCal_v2.css">
<script language="JavaScript" src="/includes/spiffyCal_v2.js"></script>
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
<form action="details.asp" method="post" name="frm" id="frm" enctype="multipart/form-data">
   <input type="hidden" name="ID" value="<%=ID%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="2" >
  <tr> 
    <td valign="top" align="left" nowrap="nowrap" colspan="2" class="titre"><h1>Camion: <%=marque & " - " & modele%></h1></td>
  </tr>
  <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="2" class="cellule1"><a href="default.asp"><img src="/manager/images/fup.gif" width="16" height="16" border="0" alt="Back to the list" /></a></th>
  </tr>
  <tr class="titlebar"> 
    <td align="left" nowrap="nowrap" colspan="2">G&eacute;n&eacute;ral</td>
  </tr>

    <tr> 
      <td nowrap="nowrap" class="cell_label">Unit&eacute; No:</td>
      <td width="100%" class="cell_content"> <input type="text" name="unite" size="40" value="<%=unite%>" class="field" />      </td>
    </tr>
    <tr>
      <td nowrap="nowrap" class="cell_label">Succursale:</td>
      <td class="cell_content">
        <input type="text" name="succursale" size="40" value="<%=succursale%>" class="field" />      </td>
    </tr>
    <tr> 
      <td nowrap="nowrap" class="cell_label">Marque:</td>
      <td class="cell_content"> <input type="text" name="marque" size="40" value="<%=marque%>" class="field" />      </td>
    </tr>
    <tr> 
      <td nowrap="nowrap" class="cell_label">Mod&egrave;le:</td>
      <td class="cell_content"> <input type="text" name="modele" size="40" value="<%=modele%>" class="field" />      </td>
    </tr>
	<tr>
		<td nowrap="nowrap" class="cell_label">Config:</td>
		<td class="cell_content"><select name="config" id="config">
            <option>-</option>
            <option <%if config="4 x 4" then response.write "selected"%>>4 x 4</option>
            <option <%if config="4 x 2" then response.write "selected"%>>4 x 2</option>
            <option <%if config="6 x 4" then response.write "selected"%>>6 x 4</option>
        </select></td>
	</tr>
    <tr> 
      <td nowrap="nowrap" class="cell_label">Ann&eacute;e:</td>
      <td class="cell_content"> <input type="text" name="intAnnee" size="40" value="<%=intAnnee%>" class="field" />      </td>
    </tr>
    <tr> 
      <td nowrap="nowrap" class="cell_label">Kilom&eacute;trage:</td>
      <td class="cell_content"> <input name="intMillage" type="text" class="field" value="<%=intMillage%>" size="40" />
         km
(Vide=N/D)      </td>
    </tr>
    <tr> 
      <td nowrap="nowrap" class="cell_label">No S&eacute;rie:</td>
      <td class="cell_content"> <input name="noSerie" type="text" class="field" value="<%=noSerie%>" size="40" /> 
		Specs (pdf): 	
        <%
	  FileExists=False
	  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	  if FSO.FileExists(Server.MapPath("/pics/specs/spec" & ID & ".pdf")) then
		%>
        <a href="/pics/specs/spec<%=ID%>.pdf" target="_blank"><img src="/intranet/images/ico_pdf.gif" alt="PDF" width="16" height="16" border="0" align="absmiddle" /></a>
        <%
		FileExists=True
	  End If
	  Set FSO = Nothing
	  
		%>
        <%IF FileExists Then%>
        <a href="#" onClick="DeleteSpec(<%=ID%>);"><img src="/intranet/images/ico_del.gif" alt="Supprimer" width="16" height="16" border="0" align="absmiddle" /></a>
        <%End If 
		%>
      
      <input type="file" name="_spec"/></td>
    </tr>
    <tr>
      <td nowrap="nowrap" class="cell_label">Date d'acquisition:</td>
      <td class="cell_content"><script language="JavaScript" type="text/javascript">
	  var cal2=new ctlSpiffyCalendarBox("cal2", "frm", "dateAchat","btnDate2","<%=MakeLocalDate(dateAchat)%>",scBTNMODE_CALBTN,2);
      cal2.writeControl();</script></td>
    </tr>
    <tr>
      <td nowrap="nowrap" class="cell_label"># de mois en inv.:</td>
      <td class="cell_content">&nbsp;<%=DateDiff("m",dateAchat,Date())+0%></td>
    </tr>
<% If ID <>"" then %>
    <tr class="titlebar">
      <td align="left" nowrap="nowrap" colspan="2">R&eacute;servation</td>
    </tr>
	<%IF resUsers_ID=0 then%>	
        <tr>
          <td nowrap="nowrap" class="cell_label">&Eacute;tat:</td>
          <td class="cell_content">&nbsp;Pr&eacute;sentement disponible</td>
        </tr>
    <%Else%>
        <tr>
          <td nowrap="nowrap" class="cell_label">&Eacute;tat:</td>
          <td class="cell_content">&nbsp;Pr&eacute;sentement r&eacute;serv&eacute; par <a href="mailto:<%=GetUserEmail(resUsers_ID)%>"><%=GetUserName(resUsers_ID)%></a> jusqu'au <%=MakeLocalDate(resDateHeure) & " " & Right(formatNumber(Hour(resDateHeure)/100,2),2) & ":" & Right(formatNumber(Minute(resDateHeure)/100,2),2)%></td>
        </tr>
    <%End If
	IF Session("Level")=1 then
		If resUsers_ID=0 then
	%>
		<tr>
		  <td align="left" nowrap="nowrap" class="cell_label">R&eacute;serv&eacute; pour:</td>
		  <td class="cell_content"><select name="_ResUsers_ID" class="field">
		  <%=MakeUserCombo(0)%>
		  </select></td>
		</tr>
		<tr>
		  <td align="left" nowrap="nowrap" class="cell_label">Jusqu'au:</td>
		  <td class="cell_content"><script language="JavaScript" type="text/javascript">
		  var cal1=new ctlSpiffyCalendarBox("cal1", "frm", "_Date","btnDate1","<%=MakeLocalDate(Now())%>",scBTNMODE_CALBTN,2);
		  cal1.writeControl();</script>
	<select name="_HH" class="field"><%=hhCombo%></select>:<select name="_MM" class="field"><%=mmCombo%></select></td>
		</tr>
		<tr>
		  <td align="left" nowrap="nowrap" class="cell_label">&nbsp;</td>
		  <td class="cell_content"><input name="_Reserve" type="submit" class="field" value="R&eacute;server!" /></td>
		</tr>
		<%Else%>
		<tr>
		  <td align="left" nowrap="nowrap" class="cell_label">&nbsp;</td>
		<td class="cell_content"><input name="_Cancel" type="submit" class="field" value="Lib&eacute;rer!" /></td>    </tr>
	<% End If
	End If%>
<%End If%>
    <tr class="titlebar"> 
      <td align="left" nowrap="nowrap" colspan="2">Photos</td>
    </tr>
    
    
    

    <tr>
      <td valign="top" class="cell_label"><label>Images :</label>
      <p style="font-weight:normal;margin:0;"><em>Utiliser votre souris pour modifier l'ordre d'affichage des images</em></p></td>
      <td class="cell_content" id="allpics"><%
                       Set rs=Conn.execute("Select * from pictures WHERE product_id='" & ID & "' ORDER BY intorder")
                       Do While Not rs.EOF
                       %>
                  <table border="0" cellspacing="0" cellpadding="0" style="float:left;margin:0 5px 5px 0;border:1px solid #666; background-color:#EEE;padding:1px;">
                    <tr>
                      <td colspan="2" align="center"><a data-lightbox="<%=rs("name")%>"  href="/pics/lg/<%=rs("name")%>" target="_blank"><img src="/pics/small/<%=rs("name")%>" border="1" width="60" height="60" alt="Click for a larger image" /></a></td>
                    </tr>
					<%IF Session("Level")=1 then%>
                    <tr>
                      <td class='pic_sort'><input name="_picture_id" type="hidden" value="<%=rs("id")%>" /></td>
                      <td align="right"><input type="button" onclick="DeletePic(<%=ID%>,<%=rs("id")%>);" value="X" alt="Remove" class="smbutton" style="width:18px;height:18px;" /></td>
                    </tr>
                    <%end IF%>
                  </table>
          <%
                            rs.MoveNext
                        Loop
                        %></td>
    </tr>
    <%IF Session("Level")=1 then%>
    <tr>
      <td valign="top" class="cell_label">&nbsp;</td>
      <td class="cell_content" id="browsetd"><input type="file" name="_img_file"/>
          <input type="button" id="addpic" value="+" class="smbutton" style="width:18px;height:18px;" />      </td>
    </tr>
	<%end if%>

    <tr class="titlebar"> 
      <td align="left" nowrap="nowrap" colspan="2">Documents</td>
    </tr>
    
    
    

    <tr>
      <td valign="top" class="cell_label"><label>Fichiers :</label></td>
      <td class="cell_content" id="alldocs"><%
                       Set rs=Conn.execute("Select * from docs WHERE product_id='" & ID & "' ORDER BY intorder")
                       Do While Not rs.EOF
                       %>
                  <table border="0" cellspacing="0" cellpadding="0" style="float:left;margin:0 5px 5px 0;border:1px solid #666; background-color:#EEE;padding:1px;">
                    <tr>
                      <td colspan="2" align="center"><a href="/pics/docs/doc<%=rs("id")%>.pdf" target="_blank"><img src="/intranet/images/ico_pdf.gif" border="0" width="32" height="32"/></a></td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><%=rs("name")%></td>
                    </tr>
					<%IF Session("Level")=1 then%>
                    <tr>
                      <td class='doc_sort'><input name="_doc_id" type="hidden" value="<%=rs("id")%>" /></td>
                      <td align="right"><input type="button" onclick="DeleteDoc(<%=ID%>,<%=rs("id")%>);" value="X" alt="Remove" class="smbutton" style="width:18px;height:18px;" /></td>
                    </tr>
                    <%end IF%>
                  </table>
          <%
                            rs.MoveNext
                        Loop
                        %></td>
    </tr>
    <%IF Session("Level")=1 then%>
    <tr>
      <td valign="top" class="cell_label">&nbsp;</td>
      <td class="cell_content" id="browsedoctd"><input type="file" name="_doc_file"/>
          <input type="button" id="adddoc" value="+" class="smbutton" style="width:18px;height:18px;" />      </td>
    </tr>
	<%end if%>


    <tr class="titlebar"> 
      <td align="left" nowrap="nowrap" colspan="2">D&eacute;tails</td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Empattement:</td>
      <td class="cell_content"> <input name="empattement" type="text" class="field" id="empattement" value="<%=empattement%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Essieu avant:</td>
      <td class="cell_content"> <input name="essieu_avant" type="text" class="field" id="essieu_avant" value="<%=essieu_avant%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Essieu arri&egrave;re:</td>
      <td class="cell_content"> <input name="essieu_arriere" type="text" class="field" id="essieu_arriere" value="<%=essieu_arriere%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Suspension arri&egrave;re:</td>
      <td class="cell_content"> <input name="suspension_ar" type="text" class="field" value="<%=suspension_ar%>" size="40" />      </td>
    </tr>
    
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Transmission:</td>
      <td class="cell_content"> <input name="transmission" type="text" class="field" value="<%=transmission%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Moteur:</td>
      <td class="cell_content"> <input name="moteur" type="text" class="field" id="moteur" value="<%=moteur%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Ratio essieu arri&egrave;re:</td>
      <td class="cell_content"> <input name="ratio_ar" type="text" class="field" id="ratio_ar" value="<%=ratio_ar%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" valign="top" nowrap="nowrap" class="cell_label">Pneus:</td>
      <td valign="top" class="cell_content"> 
		<div>
			<table border="0" cellspacing="0" cellpadding="0" style="float: left;">
			  <tr class="cell_label">
				<td>&nbsp;</td>
				<td>Dimension</td>
				<td>&Eacute;tat</td>
			  </tr>
			  <tr>
				<td class="cell_label">Avant:</td>
				<td><input name="pneu_av_dim" type="text" class="field" id="pneu_av_dim" value="<%=pneu_av_dim%>" size="25" /></td>
				<td><input name="pneu_av_etat" type="text" class="field" id="pneu_av_etat" value="<%=pneu_av_etat%>" size="25" /></td>
			  </tr>
			  <tr>
				<td class="cell_label">Arri&egrave;re:</td>
				<td><input name="pneu_ar_dim" type="text" class="field" id="pneu_ar_dim" value="<%=pneu_ar_dim%>" size="25" /></td>
				<td><input name="pneu_ar_etat" type="text" class="field" id="pneu_ar_etat" value="<%=pneu_ar_etat%>" size="25" /></td>
			  </tr>
			</table>
			<div style="float: left;">&nbsp;&nbsp;&nbsp;&nbsp;</div>
			<table border="0" cellspacing="0" cellpadding="0" style="float: left;">
				<tr class="cell_label">
					<td>X / 32</td>
					<td>Av.</td>
					<td>Av. Ar.</td>
					<td>Ar.</td>
					<td>Ar. Ar.</td>
				</tr>
				<tr>
					<td class="cell_label">Droite</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td><input name="Pneu32_ArDExt" type="text" class="field" id="Pneu32_ArDExt" value="<%=Pneu32_ArDExt%>" size="10" /></td>
					<td><input name="Pneu32_ArArDExt" type="text" class="field" id="Pneu32_ArArDExt" value="<%=Pneu32_ArArDExt%>" size="10" /></td>
				</tr>
				<tr>
					<td class="cell_label">Droite</td>
					<td><input name="Pneu32_AvD" type="text" class="field" id="Pneu32_AvD" value="<%=Pneu32_AvD%>" size="10" /></td>
					<td><input name="Pneu32_AvAvD" type="text" class="field" id="Pneu32_AvAvD" value="<%=Pneu32_AvAvD%>" size="10" /></td>
					<td><input name="Pneu32_ArDInt" type="text" class="field" id="Pneu32_ArDInt" value="<%=Pneu32_ArDInt%>" size="10" /></td>
					<td><input name="Pneu32_ArArDInt" type="text" class="field" id="Pneu32_ArArDInt" value="<%=Pneu32_ArArDInt%>" size="10" /></td>
				</tr>
				<tr>
					<td class="cell_label">Gauche</td>
					<td><input name="Pneu32_AvG" type="text" class="field" id="Pneu32_AvG" value="<%=Pneu32_AvG%>" size="10" /></td>
					<td><input name="Pneu32_AvAvG" type="text" class="field" id="Pneu32_AvAvG" value="<%=Pneu32_AvAvG%>" size="10" /></td>
					<td><input name="Pneu32_ArGInt" type="text" class="field" id="Pneu32_ArGInt" value="<%=Pneu32_ArGInt%>" size="10" /></td>
					<td><input name="Pneu32_ArArGInt" type="text" class="field" id="Pneu32_ArArGInt" value="<%=Pneu32_ArArGInt%>" size="10" /></td>
				</tr>
				<tr>
					<td class="cell_label">Gauche</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td><input name="Pneu32_ArGExt" type="text" class="field" id="Pneu32_ArGExt" value="<%=Pneu32_ArGExt%>" size="10" /></td>
					<td><input name="Pneu32_ArArGExt" type="text" class="field" id="Pneu32_ArArGExt" value="<%=Pneu32_ArArGExt%>" size="10" /></td>
				</tr>
			</table>
		</div>
	  </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Freins:</td>
      <td class="cell_content"> <input name="freins" type="text" class="field" value="<%=freins%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">R&eacute;servoir:</td>
      <td class="cell_content"> <input name="reservoirs" type="text" class="field" id="reservoirs" value="<%=reservoirs%>" size="80" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Couleur int&eacute;rieur:</td>
      <td class="cell_content"> <input name="couleur_in" type="text" class="field" id="couleur_in" value="<%=couleur_in%>" size="40" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Couleur ext&eacute;rieur:</td>
      <td class="cell_content"> <input name="couleur_ex" type="text" class="field" id="couleur_ex" value="<%=couleur_ex%>" size="40" />      </td>
    </tr>

    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label" valign="top">&Eacute;quipements:</td>
      <td class="cell_content"> <textarea name="equipements" cols="80" class="field" rows="8"><%=equipements%></textarea>            </td>
    </tr>
    <tr class="titlebar"> 
      <td align="left" nowrap="nowrap" colspan="2">Notes (usage interne seulement)</td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label" valign="top">Notes</td>
      <td class="cell_content"> <textarea name="notes" cols="80" class="field" rows="8"><%=notes%></textarea>      </td>
    </tr>
    <tr class="titlebar">
      <td align="left" nowrap="nowrap" colspan="2">Informations suppl&eacute;mentaires
        (liste imprimable)</td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">&Eacute;quipemenst (liste):</td>
      <td class="cell_content">
        <input name="equipement2" type="text" class="field" value="<%=equipement2%>" size="80" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Ancien Client:</td>
      <td class="cell_content">
        <input name="ancien_client" type="text" class="field" value="<%=ancien_client%>" size="40" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">KM Moteur:</td>
      <td class="cell_content">
        <input name="km_moteur" type="text" class="field" value="<%=km_moteur%>" size="40" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Essieux:</td>
      <td class="cell_content">
        <input name="essieux" type="text" class="field" value="<%=essieux%>" size="40" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Prix:</td>
      <td class="cell_content">
        <input name="prix" type="text" class="field" value="<%=prix%>" size="40" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Bonis:</td>
      <td class="cell_content"><input name="bonis" type="text" class="field" id="bonis" value="<%=bonis%>" size="40" />      </td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label">Promo financement:</td>
      <td class="cell_content">
        <input name="promo" type="text" class="field" id="promo" value="<%=promo%>" size="10" />
        %<em> (laisser vide pour d&eacute;sactiver)</em></td>
    </tr>

    <tr class="titlebar"> 
      <td align="left" nowrap="nowrap" colspan="2">Actions</td>
    </tr>
    <tr>
      <td align="left" nowrap="nowrap" class="cell_label" >&nbsp;</td>
      <td align="left" class="cell_content" ><table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;Fiche active</td>
            <td><input <%IF intStatus then response.Write("checked")%> type="checkbox" name="intStatus" value="1" /></td>
          </tr>
        </table></td>
    </tr>
    <%IF Session("Level")=1 then%>
    	
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label" >&nbsp;</td>
      <td align="left" class="cell_content" > <input type="submit" value="Sauvegarder la fiche" class="button" /> 
        <% If ID <> "" then%> <input name="_clone" type="submit" class="button" id="_clone" value="Dupliquer" />
        <input onclick="Delete(<%=ID%>)" type="button" value="Effacer la fiche" name="del" class="button" /> 
        <%End If%> </td>
    </tr>
    <%End IF%>
</table>

</form>
<div id="spiffycalendar" class="text">&nbsp;</div>
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
<%
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>