<div align="center">

## pjl\_setmessage


</div>

### Description

A small module to set the status message on PJL printers (HP's etc).
 
### More Info
 
Printer Name

Message

This is the code for the module


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Laurence](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/laurence.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/laurence-pjl-setmessage__1-14794/archive/master.zip)

### API Declarations

```
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DocInfo) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
```


### Source Code

```
' PJL.bas - set the status message on PJL printers (HP LaserJet etc)
'
' Based on Q154078 at support.microsoft.com which says how to write raw data to the printer
' and plint (a qbasic program)
'
Option Explicit
'
' Structure required by StartDocPrinter
'
Private Type DocInfo
 pDocName As String
 pOutputFile As String
 pDatatype As String
End Type
Dim hPrinter As Long
Dim pjlHeader As String
Dim pjlRdyMsg As String
Dim pjlFooter As String
Private Sub InitEscapeCodes()
' Private function to setup escape codes
pjlHeader = Chr(27) & "%-12345X@PJL" & vbLf
pjlRdyMsg = "@PJL RDYMSG DISPLAY="
pjlFooter = Chr(27) & "%-12345X" & vbLf
End Sub
Public Sub PJL_OpenPrinter(PrinterName As String)
' Call this function before you start sending messages
' Normally set PrinterName to Printer.DeviceName, but you might want to print to the non default printer
Dim MyDoc As DocInfo
If OpenPrinter(PrinterName, hPrinter, 0) = 0 Then MsgBox "Can't print to " & PrinterName: Exit Sub
MyDoc.pDocName = "Document"
MyDoc.pOutputFile = vbNullString
MyDoc.pDatatype = vbNullString
StartDocPrinter hPrinter, 1, MyDoc
Call StartPagePrinter(hPrinter)
InitEscapeCodes
End Sub
Public Sub PJL_ClosePrinter()
' Call this when you have finished writing messages, then they will be spooled
EndPagePrinter hPrinter
EndDocPrinter hPrinter
ClosePrinter hPrinter
hPrinter = Empty
End Sub
Public Sub PJL_WriteMessage(message As String)
' Call this to set a message for the display
' If string is too long for screen it will chop off the end
' If you have two lines on your printer the second line is just a continuation of the first
' If you set it more than once the lines will appear one after the other with 1s delay between them
Dim bDone As Long: Dim pjlCmd As String
If hPrinter = Empty Then MsgBox "Please open the printer first"
pjlCmd = pjlRdyMsg & Chr(34) & message & Chr(34) & vbLf
WritePrinter hPrinter, ByVal pjlHeader, Len(pjlHeader), bDone
WritePrinter hPrinter, ByVal pjlCmd, Len(pjlCmd), bDone
WritePrinter hPrinter, ByVal pjlFooter, Len(pjlFooter), bDone
End Sub
```

