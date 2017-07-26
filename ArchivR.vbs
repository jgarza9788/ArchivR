'objArgs is all objects that are dropped on this file
Set objArgs = Wscript.Arguments

'create a shell to run cmd
Dim objShell
Set objShell = WScript.CreateObject ("WScript.shell")

'if there is nothing dropped on this file
if objArgs.count = 0 then
    msgbox "Please darg and drop file(s) and folder(s) on this file",VBOKOnly
end if

'used for interacting with the file system
Set objFso = createobject("scripting.filesystemobject")

'iterate through all the arguments passed
For i = 0 to objArgs.count - 1
  on error resume next


    'get data about the objscts
    Path = objFso.getAbsolutePathName(objArgs(i))
    ParentFolder = objFso.getParentFolderName(objArgs(i))
    FileName = objFso.getfileName(objArgs(i))
    BaseName = objFso.getBaseName(objArgs(i))
    Extension = objFso.getExtensionName(objArgs(i))

''no longer needed
    ' Dim isFolder
    ' if Len(Extension) = 0 then
    '     isFolder = true
    ' else
    '     isFolder = false
    ' end if

    'get date time stamp
    DtTm = CStr(year(Date)) +  Right(("00" + CStr(Month(Date))),2) + CStr(Day(Date)) + right("00" +CStr(Hour(time)),2) + Right("00" + CStr(Minute(time)),2)

    'create Archive folder if needed
    if NOT objFso.FolderExists(ParentFolder & "\Archive\") then
        objFso.CreateFolder ParentFolder & "\Archive\"
    end if 

    'run cmd to zip and make a copy
    objShell.run """C:\Program Files\7-Zip\7z.exe"" a -tzip """ & ParentFolder & "\Archive\" & BaseName & "_" & DtTm & ".zip" & """ """ & Path & """"

If Err.Number <> 0 Then
    msgbox "Error:" &  err.Number & " - unable to archive file " & FileName & vbNewLine & vbNewLine & " Make sure you have 7zip insalled," & vbnewline & " and command in line 44 is correct." , vbCritical
    Err.Number = 0
    Err.Clear
End If

Next

