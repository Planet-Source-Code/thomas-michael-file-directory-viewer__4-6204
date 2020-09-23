<div align="center">

## File / Directory Viewer


</div>

### Description

This Will Display All The Files, File Size and file date of every file in the directory you specify.

To make this work, paste the code into your favorite html editor, save it and then view it.
 
### More Info
 
File System Object Be Needed :) and it is setup to look for you my documents folder at "c:\mydocu~1" but you can change the line of code to look in any directory.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thomas Michael](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-michael.md)
**Level**          |Beginner
**User Rating**    |4.5 (59 globes from 13 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thomas-michael-file-directory-viewer__4-6204/archive/master.zip)





### Source Code

```
<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<HTML>
<HEAD> <TITLE>File Viewer</TITLE> </HEAD>
<BODY>
<Table width="100%" border=1 bordercolor="#000000" align="left" cellpadding="2" cellspacing="0">
<Tr align="left" valign="top" bgcolor="#000000">
	<TD width="65%"><font color="#FFFFFF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Title</font></b></font></Td>
  <Td width="10%"><font color="#FFFFFF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Size</font></b></font></Td>
  <Td width="25%"><font color="#FFFFFF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Date</font></b></font></Td>
</Tr>
 <%
	'File System Object
	dim objFSO
	'File Object
	dim objFile
	'Folder Object
	dim objFolder
	'String To Store The Real Path
	dim sMapPath
	'Create File System Object to get list of files
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Get The path for the web page and its dir.
	'change this setting to view different directories
	sMapPath = "C:\Mydocu~1"
	'Set the object folder to the mapped path
	Set objFolder = objFSO.GetFolder(sMapPath)
	'For Each file in the folder
	For Each objFile in objFolder.Files
	%>
	<TR align="left" valign="top" bordercolor="#999999" bgcolor="#FFFFFF">
  	<TD> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><A href="<% = sMapPath & "/" & objFile.Name %>">
	<%
			'write the files name
			Response.Write objFile.Name
	%>
	</a>
	</font>
	</TD>
  <TD>
	<font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
	<%
			'We will format the file size so it looks pretty
			If objFile.Size <1024 Then
				Response.Write objFile.Size & " Bytes"
			ElseIF objFile.Size < 1048576 Then
				Response.Write Round(objFile.Size / 1024.1) & " KB"
			Else
				Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
			End If
	%>
	</font>
	</TD>
  <TD>
		<font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
			<%	'the files date
				Response.Write objFile.DateLastModified
			%>
		</font>
	</TD>
	</font>
	</TD>
	</Tr>
	<%
		Next
	%>
</Table>
</BODY>
</HTML>
```

