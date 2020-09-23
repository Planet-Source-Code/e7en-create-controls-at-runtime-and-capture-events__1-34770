<div align="center">

## Create Controls at Runtime and Capture Events


</div>

### Description

This code will show you how to easily Create Controls at runtime and Capture there events 'With-out' useing API :).

This code will show you how to add a command button and display a message box when its clicked. Please Vote and Post Comments!!
 
### More Info
 
'Here are some more control Names

'"VB.Label"

'"VB.TextBox"

'"VB.ListBox"

'"VB.ComboBox"

'"VB.OptionButton"

'"VB.CheckBox"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ï¿½e7eN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/e7en.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/e7en-create-controls-at-runtime-and-capture-events__1-34770/archive/master.zip)





### Source Code

```
Public WithEvents Command1 As CommandButton
Private Sub Command1_Click()
MsgBox "Button Pressed"
End Sub
Private Sub Form_Load()
Set Command1 = Me.Controls.Add("VB.CommandButton", "Command1", Me)
With Command1
.Visible = True
.Width = 900
.Height = 900
.Left = Me.Width / 2 - .Width / 2
.Top = Me.Height / 2 - .Height / 2
.Caption = "Test Button"
End With
End Sub
```

