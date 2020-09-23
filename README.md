<div align="center">

## Get the Windows and System directories\.


</div>

### Description

These two functions will return the location of the Windows directory (WinDir)

and the location of the System directory (SysDir).
 
### More Info
 
A boolean value indicating whether you would like a "\" character added to the

end of the file path. Thus if you pass the value true it returns "C:\WINDOWS\"

and if you pass false it returns "C:\WINDOWS".

Put the API declarations and the functions in a standard module and you should

be ready to go. The code is pretty easy to follow so I have no commented it.

If you are having trouble understanding it and would like me to come back and

add comments explaining it just leave a comment at the bottom of the page.

A string value containing the path of either the Windows directory or the System directory.

None that I know of. Leave a comment if you find one.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timothy Pew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timothy-pew.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timothy-pew-get-the-windows-and-system-directories__1-1811/archive/master.zip)

### API Declarations

```
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
```


### Source Code

```
Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String
  Dim t As String * 255
  Dim i As Long
  i = GetWindowsDirectory(t, Len(t))
  WinDir = Left(t, i)
  If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
    WinDir = WinDir & "\"
  ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
    WinDir = Left(WinDir, Len(WinDir) - 1)
  End If
End Function
Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String
  Dim t As String * 255
  Dim i As Long
  i = GetSystemDirectory(t, Len(t))
  SysDir = Left(t, i)
  If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
    SysDir = SysDir & "\"
  ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
    SysDir = Left(SysDir, Len(SysDir) - 1)
  End If
End Function
```

