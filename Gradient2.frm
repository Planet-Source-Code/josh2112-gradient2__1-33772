VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GradientMaster"
   ClientHeight    =   6615
   ClientLeft      =   4830
   ClientTop       =   3945
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Updater 
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bottom Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
      Begin Project1.ColorBar ColorBar7 
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar8 
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar9 
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin VB.Label Label9 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "G"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Middle Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      Begin Project1.ColorBar ColorBar4 
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar5 
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar6 
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin VB.Label Label6 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "G"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Top Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin Project1.ColorBar ColorBar3 
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar2 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin Project1.ColorBar ColorBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin VB.Label Label5 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "G"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox PicBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5880
      Left            =   2160
      ScaleHeight     =   5820
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "GradientMaster"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "by Joshua Foster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  PicBox.ScaleHeight = 500
  PicBox.ScaleWidth = 1
  
  ColorBar1.Color = RGB(255, 0, 0)  'set the color bars to
  ColorBar2.Color = RGB(0, 255, 0)  'the appropriate color
  ColorBar3.Color = RGB(0, 0, 255)
  ColorBar4.Color = RGB(255, 0, 0)
  ColorBar5.Color = RGB(0, 255, 0)
  ColorBar6.Color = RGB(0, 0, 255)
  ColorBar7.Color = RGB(255, 0, 0)
  ColorBar8.Color = RGB(0, 255, 0)
  ColorBar9.Color = RGB(0, 0, 255)

End Sub

Private Sub Updater_Timer()
' called every 30 ms

  Dim incR1, incG1, incB1, _
      incR2, incG2, incB2, _
      red, green, blue As Double
  
  red = ColorBar1.Value    ' starting color
  green = ColorBar2.Value
  blue = ColorBar3.Value
  
  'determine RGB increments for each line for top
  'and bottom half of box
  incR1 = (ColorBar4.Value - red) / 250
  incG1 = (ColorBar5.Value - green) / 250
  incB1 = (ColorBar6.Value - blue) / 250
  incR2 = (ColorBar7.Value - ColorBar4.Value) / 250
  incG2 = (ColorBar8.Value - ColorBar5.Value) / 250
  incB2 = (ColorBar9.Value - ColorBar6.Value) / 250
  
  For i = 0 To 500
    ' draw horizontal line at vertical position i
    PicBox.Line (0, i)-(1, i), RGB(red, green, blue)
    ' appropriately increment the color
    If i < 250 Then
      red = red + incR1
      green = green + incG1
      blue = blue + incB1
    Else
      red = red + incR2
      green = green + incG2
      blue = blue + incB2
    End If
  Next i
  
End Sub

Private Sub SaveButton_Click()
  
  fname = InputBox("Type the path and filename for" & _
                   " your image." & vbCrLf & vbCrLf & _
                   "( .bmp extension not necessary )")
  
  If fname = "" Then Exit Sub
  ' add extension if neccesary
  If Not Right(fname, 4) = ".bmp" Then fname = fname & ".bmp"
  If Not Dir(fname) = "" Then
    answer = MsgBox("File already exists.  Overwrite?", _
                     vbYesNo + vbExclamation)
    If answer = vbNo Then Call SaveButton_Click
  End If
  
  SavePicture PicBox.Image, fname

End Sub

Private Sub ExitButton_Click()
  End
End Sub
