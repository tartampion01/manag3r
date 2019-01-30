<%@EnableSessionState=False%>
<!--#include virtual="/intranet/includes/upload/_upload.asp" -->

<%
Response.Expires = -10000
Server.ScriptTimeOut = 1000

Function ResizeImage(FileName,MaxWidth,MaxHeight,Overwrite)
	'Reduce and compress large image
	
	on error resume next
	
	Dim Image
	
	Filename = Server.MapPath(Filename)

	Set Image = Server.CreateObject("GflAx.GflAx")
	
	FileName=lcase(FileName)

	Image.LoadBitmap FileName
	Image.SaveJPEGQuality = 60
	
	'ignore  X or Y modes
	IF MaxWidth=0 then MaxWidth=Image.Width
	IF MaxHeight=0 then MaxHeight=Image.Height 
	
	If Image.Height > MaxWidth  or Image.Height > MaxHeight then
		If Image.Width > MaxWidth then
			Image.Resize MaxWidth,Int(Image.Height*MaxWidth/Image.Width) 
		End IF
		If Image.Height > MaxHeight then
			Image.Resize Int(Image.Width*MaxHeight/Image.Height),MaxHeight
		End If

		If Overwrite then
			Image.SaveBitmap FileName
		Else
			LastSlashPos = InStrRev(Filename,"\")
			Image.SaveBitmap Left(Filename,LastSlashPos) & "small\" & Right(Filename,Len(Filename)-LastSlashPos)
		End If
	
	End if

	if err.number <> 0 then
		response.write "An error occured creating Image:" & err.description & "<br>"
	End If

	Set Image = Nothing
	
	'response.end 

	
End Function


'on error resume next

Set Form = New ASPForm
Form.SizeLimit = &HA00000
Form.UploadID = Request.QueryString("UploadID")
If Form.State = fsCompletted Then 'Completted
  'was the Form successfully received?
  if Form.State = 0 then
	  For Each File In Form.Files.Items
			If left(form("filename"),3)="pic" then
				File.SaveAs Server.MapPath(form("vpath") & form("filename"))
				Call ResizeImage(form("vpath") & form("filename"),300,0,True)
				Call ResizeImage(form("vpath") & form("filename"),80,0,False)
			Else
				File.SaveAs Server.MapPath(upload.form("vpath") & form("filename"))
				Call ResizeImage(Form("vpath") & form("filename"),75,75,True)
			End If
		  'File.SaveAs Server.Mappath(Form("vpath") &  form("filename") & ".pdf" )
	  Next
  End If

ElseIf Form.State > 10 then
  Const fsSizeLimit = &HD
  Select case Form.State
		case fsSizeLimit: response.write  "<br><Font Color=red>Source form size (" & Form.TotalBytes & "B) exceeds form limit (" & Form.SizeLimit & "B)</Font><br>"
		case else response.write "<br><Font Color=red>Some form error.</Font><br>"
  end Select
  
  response.end
  
End If'Form.State = 0 then


returnaddress = Form("return") & "?ID=" & Form("ID")


Response.Redirect(returnaddress)


%>