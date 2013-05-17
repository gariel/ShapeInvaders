VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invaders"
   ClientHeight    =   6195
   ClientLeft      =   5130
   ClientTop       =   2550
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8985
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   0
      Top             =   5340
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   0
      Top             =   5760
   End
   Begin VB.Label game_over 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   795
      Left            =   3000
      TabIndex        =   6
      Top             =   3180
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Shape game_over_bg 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   795
      Left            =   2880
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   8460
      TabIndex        =   4
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label bt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label bt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape shmax 
      BorderColor     =   &H80000005&
      Height          =   105
      Index           =   0
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape shmax 
      BorderColor     =   &H80000005&
      Height          =   105
      Index           =   1
      Left            =   6600
      Top             =   420
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   43
      Left            =   5700
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   42
      Left            =   6300
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   41
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   40
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   39
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   38
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   37
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   36
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   35
      Left            =   3900
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   34
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   33
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   32
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   31
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   30
      Left            =   5100
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   29
      Left            =   2100
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   28
      Left            =   1500
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   27
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   26
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   22
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   17
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   8
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   2
      Left            =   4500
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   25
      Left            =   5700
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   24
      Left            =   6300
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   23
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   21
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   20
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   19
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   18
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   16
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   15
      Left            =   3900
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   14
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   13
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   12
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   11
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   10
      Left            =   5100
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   9
      Left            =   2100
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   7
      Left            =   1500
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   6
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   5
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   4
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   3
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   1
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000D&
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Index           =   0
      Left            =   4500
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label bt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "/\"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   4260
      TabIndex        =   0
      Top             =   5460
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m As Boolean
Dim fire() As Shape
Dim ri() As Boolean
Dim rf() As Boolean
Dim ci As Long
Dim fdir As Boolean
Dim al As Integer
Dim ng As Boolean

' Ataque
Dim aX() As Integer ' X
Dim aY() As Integer ' Y
Dim a() As Integer ' Shape
Dim aP() As Long ' Attack process

Dim level As Integer
Dim v As Integer
Const vini = 15
Const level_max = 50
Const cbt = "8454016"
Const cbts = vbWhite


Private Sub bt_Click(Index As Integer)
   Select Case Index
      Case 0: ' Novo
         v = vini
         menu False
         ng = False
         game_over_bg.Visible = False
         game_over.Visible = False
         level = 0
         finaliza_level
      Case 1: ' Continuar
         menu False, True
      Case 2: ' Sair
         End
   End Select
End Sub

Private Sub bt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseMove Button, Shift, bt(Index).Left + X, bt(Index).Top + Y
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And Not ng Then menu m, True
End Sub

Private Sub Form_Load()
   m = False
   ReDim fire(0)
   ReDim ri(0)
   ReDim foe(0)
   ReDim a(0)
   ReDim aX(0)
   ReDim aY(0)
   ReDim aP(0)
   ReDim rf(0)
   fdir = True
   level = 1
   ng = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Atira
   shoot X, Label1.Top, False
End Sub

Public Sub shoot(X As Single, Y As Single, foe As Boolean)
   If Not m Then Exit Sub
   Dim i As Integer
   Dim p As Boolean
   p = False
   For i = 1 To UBound(ri)
      If Not ri(i) Then p = True: Exit For
   Next i
   
   If p Then
      fire(i).Top = Y
      fire(i).Left = X
      fire(i).Visible = True
      ri(i) = True
      rf(i) = foe
   Else
      ReDim Preserve fire(UBound(fire) + 1)
      ReDim Preserve ri(UBound(ri) + 1)
      ReDim Preserve rf(UBound(rf) + 1)
      Set fire(UBound(fire)) = Form1.Controls.Add("vb.shape", "f_" & ci, Form1)
      ci = ci + 1
      With fire(UBound(fire))
         .Width = 10
         .Height = 100
         .Top = Y
         .Left = X
         .BackColor = vbWhite
         .BorderColor = vbWhite
         .FillColor = vbWhite
         .BorderColor = vbWhite
         .FillStyle = vbDiagonalCross
         .Shape = 4 'Rounded Retangle
         .Visible = True
         .ZOrder 1
      End With
      ri(UBound(ri)) = True
      rf(UBound(rf)) = foe
   End If
End Sub

Public Function li(l As Variant, Optional delimiter As String) As String
   Dim i As Integer
   If delimiter = "" Then delimiter = "|"
   For i = 1 To UBound(l)
      li = li & delimiter & l(i)
   Next i
   li = li & delimiter
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'move
   If m Then
      If X < Form1.Width - Label1.Width And X - Label1.Width / 2 > 0 Then Label1.Left = X - Label1.Width / 2
   Else
      Dim i As Integer
      For i = 0 To bt.Count - 1
         bt(i).ForeColor = IIf(X >= bt(i).Left And X <= bt(i).Left + bt(i).Width And _
         Y >= bt(i).Top And Y <= bt(i).Top + bt(i).Height, cbts, cbt)
      Next i
   End If
   'Form_MouseDown Button, Shift, X, Y ' RISOS - atira ao mover q
End Sub

Private Sub game_over_Click()
   menu True, False
   game_over.Visible = False
   game_over_bg.Visible = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseDown Button, Shift, Label1.Left + X, Label1.Top + Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseMove Button, Shift, Label1.Left + X, Label1.Top + Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseMove Button, Shift, Label2.Left + X, Label2.Top + Y
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseDown Button, Shift, Label3.Left + X, Label3.Top + Y
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseMove Button, Shift, Label3.Left + X, Label3.Top + Y
End Sub

Private Sub Timer1_Timer()
   'loop 60 fps
   
   ' ---- Trata tiros ----
   Dim i As Integer
   Dim j As Integer
   Dim alive As Boolean
   al = 0
   
   For i = 1 To UBound(fire)
      If ri(i) Then
         If fire(i).Top > 0 Then
            fire(i).Top = fire(i).Top + IIf(rf(i), 50 + level, -100)
         Else
            fire(i).Visible = False
            ri(i) = False
         End If
      End If
   Next i
   
   ' ---- Trata inimigos ----
   If fdir And shmax(1).Left + shmax(1).Width * 3 > Me.Width Then fdir = False
   If Not fdir And shmax(0).Left <= 0 Then fdir = True
   
   For i = 0 To Shape1.Count - 1
      ' Tratamento de colisão
      If Shape1(i).Visible Then
         alive = True
         For j = 0 To UBound(fire)
            If ri(j) Then
               ' Pegou no inimigo
               If Not rf(j) And fire(j).Top >= Shape1(i).Top And fire(j).Top <= Shape1(i).Top + Shape1(i).Height And _
               fire(j).Left >= Shape1(i).Left And fire(j).Left <= Shape1(i).Left + Shape1(i).Width Then
                  Shape1(i).Visible = False
                  fire(j).Visible = False
                  ri(j) = False
                  alive = False
                  
               ' Pegou no player
               ElseIf rf(j) And fire(j).Top + fire(j).Height >= Label1.Top And fire(j).Top <= Label1.Top + Label1.Height And _
               fire(j).Left >= Label1.Left And fire(j).Left <= Label1.Left + Label1.Width Then
                  result False
                  Exit Sub
               End If
               
            End If
         Next j
      End If
      
      ' Movimenta
      If alive Then al = al + 1
      If Not flying(i) Then Shape1(i).Left = Shape1(i).Left + IIf(fdir, v, -v)
   Next i
   For j = 1 To UBound(a)
      aX(j) = aX(j) + IIf(fdir, v, -v)
   Next j
   shmax(0).Left = shmax(0).Left + IIf(fdir, v, -v)
   shmax(1).Left = shmax(1).Left + IIf(fdir, v, -v)
   If al = 0 Then finaliza_level: Exit Sub
   fly
End Sub

Public Sub finaliza_level()
   If level < level_max Then
      Dim i As Integer
      level = level + 1
      v = v + 1
      Timer2.Interval = 5100 - (1000 * level)
      Label2.Caption = level
      For i = 1 To UBound(a)
         Shape1(a(i)).Top = aY(i)
         Shape1(a(i)).Left = aX(i)
      Next i
      ReDim a(0)
      ReDim aX(0)
      ReDim aY(0)
      ReDim aP(0)
      menu False, False
   End If
End Sub

Public Sub menu(st As Boolean, Optional pause As Boolean)
   Dim i As Integer

   Timer1.Enabled = Not st
   Timer2.Enabled = Not st
   Label1.Visible = Not st
   
   Form1.MousePointer = IIf(st, 0, vbHidden)
   bt(0).Visible = st
   bt(1).Visible = st And pause
   bt(2).Visible = st
   
   If Not pause Then
      For i = 0 To Shape1.Count - 1
         Shape1(i).Visible = Not st
      Next i
      For i = 1 To UBound(ri)
         ri(i) = False
         fire(i).Visible = False
      Next i
   End If
   m = Not st
End Sub

Public Sub fly()
   'voa amiguinho
   
   Dim i As Integer
   Dim ve As Integer
   Dim ho As Integer
   Dim has_to_stop() As Integer
   ReDim has_to_stop(0)

   For i = 1 To UBound(a)
      ve = 0
      ho = 0
      With Shape1(a(i))
         If .Visible Then
            Select Case aP(i)
               Case 0 To 50 - level 'preguiça de fazer mais legal
                  ve = 3
               Case 50 - level + 1 To 8999 ' só vai
                  ve = 3 * level ' com level 50 fica cabreirão
                  ho = 3 * level
               Case 9000 To 99999 ' voltando pro lugar de origem
                  If aX(i) < .Left Then .Left = IIf(.Left - aX(i) < v * 3, aX(i), .Left - v * 3)
                  If aX(i) > .Left Then .Left = IIf(aX(i) - .Left < v * 3, aX(i), .Left + v * 3)
                  If aY(i) > .Top Then .Top = .Top + IIf(aY(i) - .Top < v * 2, aY(i) - .Top, v * 2)
                  If aY(i) = .Top And aX(i) = .Left Then
                     ReDim has_to_stop(UBound(has_to_stop) + 1)
                     has_to_stop(UBound(has_to_stop)) = i
                  End If
            End Select
            .Top = .Top + ve
            .Left = .Left + IIf(IIf(.Left + .Width / 2 > Label1.Left + Label1.Width / 2, 0, 1), ho, -ho)
            
            ' Atira
            If aP(i) Mod 100 = 0 Then shoot .Left + .Width / 2, .Top + .Height, True
            
            ' --- Ve se acertou o player ---
            
            If .Left + .Width > Label1.Left And .Left < Label1.Left + Label1.Width And _
            .Top + .Height > Label1.Top And .Top < Label1.Top + Label1.Height Then result False: Exit Sub
            
            '-------------------------------
            
            If .Top > Me.Height Then
               .Top = -.Height
               aP(i) = 9000 'Finalizacao
            End If
            aP(i) = aP(i) + 1
         End If
      End With
   Next i
   For i = 1 To UBound(has_to_stop)
      stop_attack has_to_stop(i)
   Next i
End Sub

Public Sub result(r As Boolean)
   menu True, False
   game_over.Visible = Not r
   game_over_bg.Visible = Not r
   Timer1.Enabled = False
   Timer2.Enabled = False
   ng = True
End Sub

Public Sub stop_attack(i As Integer)
   ' ver uma forma melhor de fazer sacoisa
   Dim j As Integer
   Dim n As Integer
   Dim a2() As Integer
   Dim ay2() As Integer
   Dim ax2() As Integer
   Dim ap2() As Long
   n = UBound(a) - 1
   ReDim a2(n)
   ReDim ay2(n)
   ReDim ax2(n)
   ReDim ap2(n)
   
   For j = 0 To UBound(a)
      If j <> i Then
         n = IIf(i > j, 0, 1)
         a2(j - n) = a(j)
         ay2(j - n) = aY(j)
         ax2(j - n) = aX(j)
         ap2(j - n) = aP(j)
      End If
   Next j
   a = a2
   aY = ay2
   aX = ax2
   aP = ap2
End Sub

Private Function flying(i As Integer) As Boolean
   Dim f As String
   f = li(a)
   flying = InStr(f, "|" & i & "|") > 0
End Function

Private Sub Timer2_Timer()
   ' Inicia ataque \o/
   Randomize Timer
   Dim n As Integer
   Dim i As Integer
   
   Dim f As Boolean
   Dim cc As Boolean
   f = True
   
   
   If Timer2.Interval > 300 Then Timer2.Interval = Timer2.Interval * 0.9 - (level / 100)
   n = Int((al - UBound(a) + 1) * Rnd) + 1
   i = 0
   Label3.Caption = n
   Do While f And UBound(a) < Shape1.Count - 1
      If Shape1(i).Visible And Not flying(i) Then
         cc = True
         If n < 0 Then n = n * -1 ' gato
         n = n - 1
         If n = 0 Then
            Dim p As Integer
            p = UBound(a) + 1
            ReDim Preserve a(p)
            ReDim Preserve aX(p)
            ReDim Preserve aY(p)
            ReDim Preserve aP(p)
            a(p) = i
            aX(p) = Shape1(i).Left
            aY(p) = Shape1(i).Top
            aP(p) = 0
            f = False
         End If
      End If
      If i >= Shape1.Count - 1 Then
         If Not cc Then Exit Do
         i = -1
         cc = False
      End If
      i = i + 1
   Loop
End Sub
