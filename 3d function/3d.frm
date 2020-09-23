VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3d rotation"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   900
      Top             =   4995
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Text            =   "t/15 "
      Top             =   3870
      Width           =   2490
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Text            =   " cos(t)"
      Top             =   3435
      Width           =   2490
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Text            =   " sin(t)"
      Top             =   3015
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6075
      Left            =   2520
      ScaleHeight     =   6015
      ScaleWidth      =   6285
      TabIndex        =   1
      Top             =   -45
      Width           =   6345
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "rotate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1215
      Width           =   2490
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080FFFF&
      Caption         =   "z axis"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   750
      Width           =   2580
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FFFF&
      Caption         =   "y axis"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "x axis"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2625
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2565
      Top             =   4500
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   45
      Max             =   20
      Min             =   1
      TabIndex        =   2
      Top             =   2205
      Value           =   1
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "draw"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4500
      Width           =   2445
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   " functions "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   2565
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   1755
      Width           =   2580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clr As Long
Dim alph As Double
Dim shiftx() As Double, shifty() As Double, shiftz() As Double
Dim px() As Double, py() As Double, pz() As Double 'points matrix
Dim n As Long
Const cm = 567, pi = 22 / 7
'parametric function these gives equation of helix

Private Sub Command1_Click() 'draw button
Command3.Enabled = True
alph = 0
w = ScaleWidth / 2: h = ScaleHeight / 2
l1 = -10: l2 = 10
readpoints l1, l2
 draw
End Sub
Sub coord() 'draw coordinates
w = Picture1.Width / 2: h = Picture1.Height / 2
Form1.DrawWidth = 1
l = 8
X1 = l: Y1 = 0: z1 = 0
xs1 = (((3) ^ (1 / 2)) / 2) * X1 - (((3) ^ (1 / 2)) / 2) * Y1
ys1 = z1 - Y1 / 2 - X1 / 2
Picture1.Line (w, h)-(w + xs1 * cm, h - ys1 * cm), QBColor(9)
X1 = 0: Y1 = l: z1 = 0
xs1 = (((3) ^ (1 / 2)) / 2) * X1 - (((3) ^ (1 / 2)) / 2) * Y1
ys1 = z1 - Y1 / 2 - X1 / 2
Picture1.Line (w, h)-(w + xs1 * cm, h - ys1 * cm), QBColor(9)
X1 = 0: Y1 = 0: z1 = l
xs1 = (((3) ^ (1 / 2)) / 2) * X1 - (((3) ^ (1 / 2)) / 2) * Y1
ys1 = z1 - Y1 / 2 - X1 / 2
Picture1.Line (w, h)-(w + xs1 * cm, h - ys1 * cm), QBColor(9)
End Sub
Sub roundx(i, alph) 'updates of x,y,z
r = (py(i) ^ 2 + pz(i) ^ 2) ^ (1 / 2)
px(i) = px(i)
py(i) = r * Cos(alph + shiftx(i))
pz(i) = r * Sin(alph + shiftx(i))
End Sub
Sub roundy(i, alph)
r = (px(i) ^ 2 + pz(i) ^ 2) ^ (1 / 2)
py(i) = py(i)
px(i) = r * Cos(alph + shifty(i))
pz(i) = r * Sin(alph + shifty(i))
End Sub
Sub roundz(i, alph)
r = (py(i) ^ 2 + px(i) ^ 2) ^ (1 / 2)
pz(i) = pz(i)
px(i) = r * Cos(alph + shiftz(i))
py(i) = r * Sin(alph + shiftz(i))
End Sub
'shifts calculate initial angles of each point
Sub shifts()
ReDim Preserve shiftx(n) As Double
ReDim Preserve shifty(n) As Double
ReDim Preserve shiftz(n) As Double
For i = 0 To n - 1
If Round(py(i), 3) = 0 Then py(i) = 0.01 'avoid division by zero
If Round(px(i), 3) = 0 Then px(i) = 0.01
'rotation round x
If Sgn(pz(i)) = -1 And Sgn(py(i)) = -1 Then shiftx(i) = pi + Atn(pz(i) / py(i)): GoTo e1:
If Sgn(pz(i)) = 1 And Sgn(py(i)) = -1 Then shiftx(i) = pi + Atn(pz(i) / py(i)): GoTo e1:
shiftx(i) = Atn(pz(i) / py(i))
e1:
'rotation round y
If Sgn(pz(i)) = -1 And Sgn(px(i)) = -1 Then shifty(i) = pi + Atn(pz(i) / px(i)): GoTo e2:
If Sgn(pz(i)) = 1 And Sgn(px(i)) = -1 Then shifty(i) = pi + Atn(pz(i) / px(i)): GoTo e2:
shifty(i) = Atn(pz(i) / px(i))
e2:
'rotation round z
If Sgn(py(i)) = -1 And Sgn(px(i)) = -1 Then shiftz(i) = pi + Atn(py(i) / px(i)): GoTo e3:
If Sgn(py(i)) = 1 And Sgn(px(i)) = -1 Then shiftz(i) = pi + Atn(py(i) / px(i)): GoTo e3:
shiftz(i) = Atn(py(i) / px(i))
e3:
Next i
End Sub
Sub readpoints(l1, l2) 'read points

n = 0
s = 0.2
'read functions from the texts
sc.AddCode "function x(t)" + vbCrLf + "x=" + Text1 + vbCrLf + "end function"
sc.AddCode "function y(t)" + vbCrLf + "y=" + Text2 + vbCrLf + "end function"
sc.AddCode "function z(t)" + vbCrLf + "z=" + Text3 + vbCrLf + "end function"

For t = l1 To l2 Step s
ReDim Preserve px(n) As Double
ReDim Preserve py(n) As Double
ReDim Preserve pz(n) As Double
px(n) = sc.Run("X", t)
 py(n) = sc.Run("Y", t)
  pz(n) = sc.Run("z", t)
  
n = n + 1
Next t
End Sub
Sub draw() 'draw
Picture1.Cls
coord
w = Picture1.Width / 2: h = Picture1.Height / 2
Form1.DrawWidth = HScroll1.Value
clr = QBColor(14)
For j = 0 To n - 2
xs = (((3) ^ (1 / 2)) / 2) * px(j) - (((3) ^ (1 / 2)) / 2) * py(j)
ys = pz(j) - py(j) / 2 - px(j) / 2
xs1 = (((3) ^ (1 / 2)) / 2) * px(j + 1) - (((3) ^ (1 / 2)) / 2) * py(j + 1)
ys1 = pz(j + 1) - py(j + 1) / 2 - px(j + 1) / 2
Picture1.Line (w + xs * cm, h - ys * cm)-(w + xs1 * cm, h - ys1 * cm), clr
Next j
End Sub

Private Sub Command3_Click() 'rotate/stop

If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "select rotation axis", vbExclamation

Exit Sub

End If

If Command3.Caption = "rotate" Then
Command3.Caption = "stop"
Timer1.Enabled = True
ElseIf Command3.Caption = "stop" Then
Command3.Caption = "rotate"
Timer1.Enabled = False
alph = 0
End If


End Sub
Private Sub Form_Load()
Timer1.Enabled = False
Command3.Enabled = False
l1 = -10: l2 = 10 'limits
readpoints l1, l2
End Sub

'to avoid the problem of transition for the option buttons
Private Sub Option1_Click()
alph = 0
End Sub
Private Sub Option2_Click()
alph = 0
End Sub
Private Sub Option3_Click()
alph = 0
End Sub
Private Sub Timer1_Timer()
'shifts calculated for first time only or for transition of rotation axis
If alph = 0 Then 'for first run
shifts
End If
If Option1.Value = True Then 'round x
For i = 0 To n - 1
roundx i, alph
Next i
ElseIf Option2.Value = True Then 'round y
For i = 0 To n - 1
roundy i, alph
 Next i
ElseIf Option3.Value = True Then 'round z
For i = 0 To n - 1

 roundz i, alph
 
Next i
End If
draw
Step = HScroll1.Value / 10

alph = alph + Step




End Sub
