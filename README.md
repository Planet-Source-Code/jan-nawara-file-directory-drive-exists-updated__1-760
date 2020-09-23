<div align="center">

## File/Directory/Drive Exists \(Updated\)


</div>

### Description

This code returns a true/false if a specified drive/directory/pathname exists.

This is a small, fast routine.
 
### More Info
 
A string containing a pathname must be passed.

If checking for a directory you must also set the second optional argument to True.

'To check if a specific drive letter exists, use strings for the PathName argument that look like this (the strings themselves should not include quotation marks):

'

'"c:"

'"c:\"

'

'Eg. DriveStat= File_Exists("c:\")

'(NOTE: The backslash is optional.)

'

'To check if a specific directory exists, use strings for the PathName argument that look like this (the strings themselves should not include quotation marks). ALSO, you must use True for the second optional argument, otherwise the function will not work on all directories.:

'

'"c:\temp\"

'"c:\windows\"

'

'Eg. DirStat = File_Exists("c:\temp", True)

'

'To check if a specific file exists, use strings for the PathName argument that look like this (the strings themselves should not include quotation marks):

'

'"c:\temp\somefile.exe"

'"c:\windows\notepad.exe"

'

'Eg. FileStat = File_Exists("c:\windows\win.ini")

True if the pathname and/or file exists.

Otherwise it returns false.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Nawara](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-nawara.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-nawara-file-directory-drive-exists-updated__1-760/archive/master.zip)





### Source Code

```
Function File_Exists(ByVal PathName As String, Optional Directory As Boolean) As Boolean
 'Returns True if the passed pathname exist
 'Otherwise returns False
 If PathName <> "" Then
 If IsMissing(Directory) Or Directory = False Then
  File_Exists = (Dir$(PathName) <> "")
 Else
  File_Exists = (Dir$(PathName, vbDirectory) <> "")
 End If
 End If
End Function
```

