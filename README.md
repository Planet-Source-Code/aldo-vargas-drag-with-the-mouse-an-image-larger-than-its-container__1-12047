<div align="center">

## Drag with the Mouse an Image Larger than Its Container


</div>

### Description

This code shows how to scroll with the mouse a large image that is contained in a small container.
 
### More Info
 
This example needs that you place a PictureBox and an Image in a form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aldo Vargas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aldo-vargas.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aldo-vargas-drag-with-the-mouse-an-image-larger-than-its-container__1-12047/archive/master.zip)





### Source Code

```
Option Explicit
Dim px As Long, py As Long
Dim gapx As Long, gapy As Long
Private Sub Form_Load()
 Set Image1.Container = Picture1
 Image1.Stretch = True
 Image1.Picture = LoadPicture("C:\Windows\Bubbles.bmp")
 Picture1.Move 60, 60, 6000, 4000
 Image1.Move -1000, -1000, 10000, 10000
 Me.Move Screen.Width \ 2 - 3100, Screen.Height \ 2 - 2250, 6200, 4500
End Sub
Private Sub image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 px = X
 py = Y
 gapx = Picture1.Width - Image1.Width
 gapy = Picture1.Height - Image1.Height
 Image1.MousePointer = 15
End Sub
Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim deltax As Long, deltay As Long
 If Button = vbLeftButton Then
  X = CLng(X)
  Y = CLng(Y)
  If Abs(X - px) < 30 Then
  ElseIf X < px Then
   deltax = Abs(X - px)
   If Image1.Left - deltax >= gapx Then
    Image1.Left = Image1.Left - deltax
   ElseIf gapx <= 0 Then
    Image1.Left = gapx
   Else
    Image1.Left = 0
   End If
   px = X + deltax
  ElseIf X > px Then
   deltax = Abs(X - px)
   If Image1.Left + deltax <= 0 Then
    Image1.Left = Image1.Left + deltax
   Else
    Image1.Left = 0
   End If
   px = X - deltax
  End If
  If Abs(Y - py) < 30 Then
  ElseIf Y < py Then
   deltay = Abs(Y - py)
   If Image1.Top - deltay >= gapy Then
    Image1.Top = Image1.Top - deltay
   ElseIf gapy <= 0 Then
    Image1.Top = gapy
   Else
    Image1.Top = 0
   End If
   py = Y + deltay
  ElseIf Y > py Then
   deltay = Abs(Y - py)
   If Image1.Top + deltay <= 0 Then
    Image1.Top = Image1.Top + deltay
   Else
    Image1.Top = 0
   End If
   py = Y - deltay
  End If
 End If
End Sub
Private Sub image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1.MousePointer = 0
End Sub
```

