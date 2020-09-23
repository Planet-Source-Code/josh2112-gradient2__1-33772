VERSION 5.00
Begin VB.UserControl ColorBar 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   ScaleHeight     =   255
   ScaleWidth      =   1095
   Begin VB.PictureBox Bar 
      Height          =   255
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "ColorBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mousedown As Boolean
Dim num, prevnum, col As Long

Private Sub UserControl_Initialize()
  
  Bar.ScaleHeight = 20
  Bar.ScaleWidth = 255
  Bar.AutoRedraw = True
  Bar.BackColor = vbBlack
  mousedown = False
  num = prevnum = 0
  
End Sub

Public Property Let Color(ByVal c As Long)
  col = c
End Property

Public Property Get Value() As Long
  
  If num < 0 Then num = 0
  If num > 255 Then num = 255
  Value = num

End Property

Private Sub UpdateBar()
  
  If num > prevnum Then
  ' fill last increment with color
    For i = 0 To 20
      Bar.Line (prevnum, i)-(num, i), col
    Next i
  Else
  ' fill last increment with black
    For i = 0 To 20
      Bar.Line (prevnum, i)-(num, i), RGB(0, 0, 0)
    Next i
  End If
  
End Sub

Private Sub Bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  mousedown = True
  prevnum = num
  num = X
  Call UpdateBar

End Sub

Private Sub Bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Not mousedown Then Exit Sub
  prevnum = num
  num = X
  Call UpdateBar

End Sub

Private Sub Bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  mousedown = False
  prevnum = num
  num = X
  Call UpdateBar

End Sub
