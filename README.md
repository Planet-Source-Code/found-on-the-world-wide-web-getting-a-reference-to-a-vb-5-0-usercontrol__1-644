<div align="center">

## Getting a Reference to a VB 5\.0 UserControl


</div>

### Description

Visual Basic 5.0 allows you to use UserControls to create ActiveX controls in your projects. The following code snippet does two things: It gets a reference to the form in which a UserControl is placed, and it gets a reference to that control on the form. by David Mendlen
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-getting-a-reference-to-a-vb-5-0-usercontrol__1-644/archive/master.zip)





### Source Code

```

Dim PControl As Object
Dim MyControl As Control
Dim AControl As Object
'Get my UserControl
For Each AControl In ParentControls
  If AControl.Name = Ambient.DisplayName Then
    Set MyControl = AControl
    Exit For
  End If
Next
'Get the Form UserControl is on
Set PControl = ParentControls.Item(1).Parent
While Not (TypeOf PControl Is Form)   Set PControl = PControl.Parent
Wend
```

