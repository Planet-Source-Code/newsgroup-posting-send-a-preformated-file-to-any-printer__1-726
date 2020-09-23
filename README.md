<div align="center">

## Send a preformated file to any printer


</div>

### Description

Q. How can I send a preformated file to a printer "as is". If I use Printer.Print then things like ESC get converted to a box or whatever chr$(27) is in the current font.

A.I'm using following code to send AutoCAD .plt-files to my printer, and it works ok for me. "Soren Staun Jorgensen" <ssj@post2.tele.dk>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-send-a-preformated-file-to-any-printer__1-726/archive/master.zip)

### API Declarations

```
Public Declare Function CopyFile& Lib "kernel32" Alias "CopyFileA" (ByVal
lpExistingFileName As String, ByVal lpNewFileName As String, ByVal
bFailIfExists As Long)
```


### Source Code

```
Public Sub SendFileToPrinter()
  Dim FileName As String
  Dim s As Long
  Dim i As Integer
  For i = 0 To frmMain.List.ListCount - 1
    If frmMain.List.Selected(i) Then
      FileName = CurFolder & "\" & frmFileList.File.List(i)
      s = SendToPort(FileName, CurPrnPort, vbNull)
      frmMain.List.Selected(i) = False
    End If
  Next i
End Sub
Public Function SendToPort(sFileName$, sPortName$, lPltFailed&)
Dim s As Long
  s = CopyFile(sFileName, sPortName, lPltFailed)
End Function
```

