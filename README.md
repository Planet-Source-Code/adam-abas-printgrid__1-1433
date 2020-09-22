<div align="center">

## PrintGrid


</div>

### Description

To Print DataBase Grid Control as a quick report

without buying an expensive tools.
 
### More Info
 
Data Base Control, Data Base Grid Control

How he,she use a data & dbgrid control in VB application.

DBGrid Record source as a Text File with adjusted Columns

Width.

No Idea, only when you increase number of columns the

processing is be slower.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Abas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-abas.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[DDE](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/dde__1-28.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-abas-printgrid__1-1433/archive/master.zip)

### API Declarations

```
	Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```


### Source Code

```
Please,If you do any changes let me know as a feed back.
If you like to have a .OCX as an Activex Email me, free no charge
but no source code.
Private Function PrintGd(ByVal GridToPrint As DBGrid, ByVal MyRecordset As Recordset) As Long
Dim x, v, b
Dim Putit As String
Dim Myrec
Dim MyField
Dim TCapion
Dim Mydash
 Screen.MousePointer = 11
 Open "C:\Printed.txt" For Output As #2
 Putit = ""
 Mydash = "-"
 For b = 0 To GridToPrint.Columns.Count - 1
  Myrec = ""
  MyField = ""
  x = GridToPrint.Columns(b).Width
  x = x / 100
  For v = 1 To x
  Mydash = Mydash + "-"
   If Mid(GridToPrint.Columns(b).Caption, v, 1) = "" Then
    Myrec = Chr(32)
   Else
    Myrec = Mid(GridToPrint.Columns(b).Caption, v, 1)
   End If
    MyField = MyField & Myrec
  Next v
   Putit = Putit & Chr(9) & MyField
   DoEvents
 '
 Next b
 Print #2, " No" & Putit
 Print #2, Mydash
Close #2
Dim Colcap
Dim Toprint
Open "C:\Printed.txt" For Append As #1
MyRecordset.MoveFirst
Dim Nox
Do While Not MyRecordset.EOF
Putit = ""
Nox = Nox + 1
For b = 0 To GridToPrint.Columns.Count - 1
If GridToPrint.Columns(b).Visible = True Then
  Myrec = ""
  MyField = ""
  x = GridToPrint.Columns(b).Width
  x = x / 100
  For v = 1 To x
  DoEvents
   If Mid(GridToPrint.Columns(b).Text, v, 1) = "" Then
    Myrec = Chr(32) 'x
   Else
    Myrec = Mid(GridToPrint.Columns(b).Text, v, 1)
   End If
   MyField = MyField & Myrec
  Next v
  DoEvents
  Putit = Putit & Chr(9) & MyField
 Else
 End If
 Next b
 Print #1, Format(Nox, "@@@") & Putit
MyRecordset.MoveNext
Loop
Close #1
Me.Refresh
Dim RetVal As Long
  RetVal = ShellExecute(Me.hwnd, _
   vbNullString, "C:\Printed.Txt", vbNullString, "c:\", SW_SHOWNORMAL)
Screen.MousePointer = 0
End Function
Private Sub Command1_Click()
Dim x
x = PrintGd(DBGrid1, Data1.Recordset)
End Sub
```

