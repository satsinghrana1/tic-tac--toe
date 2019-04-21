VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Zero-Cross Ver 1.0.0                  [ Developed By: Sat Singh  ]"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20160
   FillColor       =   &H00404040&
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":122FA
   ScaleHeight     =   11520
   ScaleWidth      =   20160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Picture         =   "Form1.frx":299EC
      ScaleHeight     =   1575
      ScaleWidth      =   11655
      TabIndex        =   4
      Top             =   0
      Width           =   11655
      Begin VB.Label pname 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   930
         Left            =   6090
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   12840
      Top             =   4680
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   2520
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   11640
      TabIndex        =   14
      Text            =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15840
      MouseIcon       =   "Form1.frx":410DE
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":413E8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox p2s 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17160
      TabIndex        =   7
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox p1s 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15915
      TabIndex        =   6
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Default         =   -1  'True
      Height          =   615
      Left            =   5640
      Picture         =   "Form1.frx":4514E
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6810
      Left            =   2820
      Picture         =   "Form1.frx":48049
      ScaleHeight     =   6810
      ScaleWidth      =   6840
      TabIndex        =   3
      Top             =   1560
      Width           =   6845
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14520
      Top             =   6360
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14040
      Top             =   6360
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13560
      Top             =   6360
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13080
      Top             =   6360
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12600
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12120
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11640
      Top             =   6360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11160
      Top             =   6360
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17040
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   2880
      Picture         =   "Form1.frx":57C1C
      ScaleHeight     =   6735
      ScaleWidth      =   6735
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Line l1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   0
         X2              =   15
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line l2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   960
         X2              =   975
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line l3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   3360
         X2              =   3375
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line l4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   5760
         X2              =   5775
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line l5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   6720
         X2              =   6735
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line l6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   0
         X2              =   15
         Y1              =   960
         Y2              =   975
      End
      Begin VB.Line l7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   0
         X2              =   15
         Y1              =   3360
         Y2              =   3375
      End
      Begin VB.Line l8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   0
         X2              =   15
         Y1              =   5760
         Y2              =   5775
      End
      Begin VB.Image i4 
         Height          =   1935
         Left            =   0
         MouseIcon       =   "Form1.frx":630E5
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i1 
         Height          =   1935
         Left            =   0
         MouseIcon       =   "Form1.frx":63527
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i2 
         Height          =   1935
         Left            =   2400
         MouseIcon       =   "Form1.frx":63969
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image i3 
         Height          =   1935
         Left            =   4800
         MouseIcon       =   "Form1.frx":63DAB
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i5 
         Height          =   1935
         Left            =   2400
         MouseIcon       =   "Form1.frx":641ED
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i6 
         Height          =   1935
         Left            =   4800
         MouseIcon       =   "Form1.frx":6462F
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i7 
         Height          =   1935
         Left            =   0
         MouseIcon       =   "Form1.frx":64A71
         Stretch         =   -1  'True
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i8 
         Height          =   1935
         Left            =   2400
         MouseIcon       =   "Form1.frx":64EB3
         Stretch         =   -1  'True
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image i9 
         Height          =   1935
         Left            =   4800
         MouseIcon       =   "Form1.frx":652F5
         Stretch         =   -1  'True
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   930
      Left            =   17460
      TabIndex        =   18
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   930
      Left            =   16260
      TabIndex        =   17
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play Again"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5160
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   16680
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   15480
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2 Score"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   17040
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1 Score"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   15840
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15720
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   15600
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   17400
      TabIndex        =   5
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   12000
      Top             =   3960
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image blank 
      Height          =   1935
      Left            =   14880
      Picture         =   "Form1.frx":65737
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image x 
      Height          =   1935
      Left            =   10440
      MouseIcon       =   "Form1.frx":67ADD
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":67DE7
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image z 
      Height          =   1815
      Left            =   12600
      Picture         =   "Form1.frx":6C4E6
      Stretch         =   -1  'True
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   15480
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_Click()
Label8.Caption = Time
Command6.Enabled = True
Timer10.Enabled = True
Timer11.Enabled = True
pname.Visible = True
Image1.Visible = True
Picture1.Visible = True
i1.Visible = True
i2.Visible = True
i3.Visible = True
i4.Visible = True
i5.Visible = True
i6.Visible = True
i7.Visible = True
i8.Visible = True
i9.Visible = True
Text1.Text = "x"
End Sub


Private Sub Command6_Click()
If p1s.Text > p2s.Text Then MsgBox "Player 1 Wins This Game With " + p1s.Text + " Score", vbInformation, "Final Result"
If p1s.Text < p2s.Text Then MsgBox "Player 2 Wins This Game With " + p2s.Text + " Score", vbInformation, "Final Result"
If p1s.Text = "0" And p2s.Text = "0" Then MsgBox "Play atleast one Game For Final Result", vbInformation, "Final Result"
If p1s.Text = p2s.Text Then
If p1s.Text <> "0" And p2s.Text <> "0" Then
MsgBox "This Game is Draw", vbInformation, "Final Result"
End If
End If
p1s.Text = "0"
p2s.Text = "0"

End Sub

Private Sub Form_Load()
Image1.Picture = blank.Picture
i1.Enabled = True
i2.Enabled = True
i3.Enabled = True
i4.Enabled = True
i5.Enabled = True
i6.Enabled = True
i7.Enabled = True
i8.Enabled = True
i9.Enabled = True
i1.Picture = blank.Picture
i2.Picture = blank.Picture
i3.Picture = blank.Picture
i4.Picture = blank.Picture
i5.Picture = blank.Picture
i6.Picture = blank.Picture
i7.Picture = blank.Picture
i8.Picture = blank.Picture
i9.Picture = blank.Picture
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
l1.X1 = 0
l1.X2 = 0
l1.Y1 = 0
l1.Y2 = 0
l1.Visible = False
l2.Y2 = 0
l2.Visible = False
l3.Y2 = 0
l3.Visible = False
l4.Y2 = 0
l4.Visible = False
l5.X1 = 6735
l5.X2 = 6735
l5.Visible = False
l6.X2 = 0
l6.Visible = False
l7.X2 = 0
l7.Visible = False
l8.X2 = 0
l8.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = &HC000&
Label3.BackColor = &H0&
Label3.BorderStyle = (0)

Label1.FontUnderline = False
Label1.ForeColor = &H80000005
End Sub

Private Sub i1_Click()

If Text1.Text = "x" Then
i1.Picture = x.Picture
Text1.Text = "0"
i1.Enabled = False
Else
If Text1.Text = "0" Then
i1.Picture = z.Picture
Text1.Text = "x"
i1.Enabled = False
End If
End If
End Sub

Private Sub i2_Click()

If Text1.Text = "x" Then
i2.Picture = x.Picture
Text1.Text = "0"
i2.Enabled = False
Else
If Text1.Text = "0" Then
i2.Picture = z.Picture
Text1.Text = "x"
i2.Enabled = False
End If
End If
End Sub

Private Sub i3_Click()

If Text1.Text = "x" Then
i3.Picture = x.Picture
Text1.Text = "0"
i3.Enabled = False
Else
If Text1.Text = "0" Then
i3.Picture = z.Picture
Text1.Text = "x"
i3.Enabled = False
End If
End If

End Sub

Private Sub i4_Click()

If Text1.Text = "x" Then
i4.Picture = x.Picture
Text1.Text = "0"
i4.Enabled = False
Else
If Text1.Text = "0" Then
i4.Picture = z.Picture
Text1.Text = "x"
i4.Enabled = False
End If
End If

End Sub

Private Sub i5_Click()

If Text1.Text = "x" Then
i5.Picture = x.Picture
Text1.Text = "0"
i5.Enabled = False
Else
If Text1.Text = "0" Then
i5.Picture = z.Picture
Text1.Text = "x"
i5.Enabled = False
End If
End If

End Sub

Private Sub i6_Click()

If Text1.Text = "x" Then
i6.Picture = x.Picture
Text1.Text = "0"
i6.Enabled = False
Else
If Text1.Text = "0" Then
i6.Picture = z.Picture
Text1.Text = "x"
i6.Enabled = False
End If
End If

End Sub

Private Sub i7_Click()

If Text1.Text = "x" Then
i7.Picture = x.Picture
Text1.Text = "0"
i7.Enabled = False
Else
If Text1.Text = "0" Then
i7.Picture = z.Picture
Text1.Text = "x"
i7.Enabled = False
End If
End If

End Sub

Private Sub i8_Click()

If Text1.Text = "x" Then
i8.Picture = x.Picture
Text1.Text = "0"
i8.Enabled = False
Else
If Text1.Text = "0" Then
i8.Picture = z.Picture
Text1.Text = "x"
i8.Enabled = False
End If
End If

End Sub

Private Sub i9_Click()

If Text1.Text = "x" Then
i9.Picture = x.Picture
Text1.Text = "0"
i9.Enabled = False
Else
If Text1.Text = "0" Then
i9.Picture = z.Picture
Text1.Text = "x"
i9.Enabled = False
End If
End If
End Sub


Private Sub Label1_Click()
Label1.Visible = False
i1.Enabled = True
i2.Enabled = True
i3.Enabled = True
i4.Enabled = True
i5.Enabled = True
i6.Enabled = True
i7.Enabled = True
i8.Enabled = True
i9.Enabled = True
i1.Picture = blank.Picture
i2.Picture = blank.Picture
i3.Picture = blank.Picture
i4.Picture = blank.Picture
i5.Picture = blank.Picture
i6.Picture = blank.Picture
i7.Picture = blank.Picture
i8.Picture = blank.Picture
i9.Picture = blank.Picture
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
l1.X1 = 0
l1.X2 = 0
l1.Y1 = 0
l1.Y2 = 0
l1.Visible = False
l2.Y2 = 0
l2.Visible = False
l3.Y2 = 0
l3.Visible = False
l4.Y2 = 0
l4.Visible = False
l5.X1 = 6735
l5.X2 = 6735
l5.Y1 = 0
l5.Visible = False
l6.X2 = 0
l6.Visible = False
l7.X2 = 0
l7.Visible = False
l8.X2 = 0
l8.Visible = False
Text1.Text = "0"
Text1.Text = "x"
Image1.Visible = True
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = True
Label1.ForeColor = &HF866&
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.BorderStyle = (1)
Label3.ForeColor = &H0&
Label3.BackColor = &HC000&

End Sub









Private Sub p1s_Change()
Label2.Caption = p1s.Text
End Sub

Private Sub p2s_Change()
Label9.Caption = p2s.Text
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = &HC000&
Label3.BackColor = &H0&
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.ForeColor = &H80000005

End Sub

Private Sub pname_Change()
If pname.Caption = "Game Complete Player 1 Wins" Then p1s.Text = p1s.Text + 1
If pname.Caption = "Game Complete Player 2 Wins" Then p2s.Text = p2s.Text + 1
End Sub

Private Sub Text1_Change()
If Text1.Text = "x" Then
pname.Caption = "Player 1 turn"
Else
If Text1.Text = "0" Then
pname.Caption = "Player 2 turn"
End If
End If
End Sub

Private Sub Text2_Change()
Label8.Caption = Text2.Text
End Sub


Private Sub Timer1_Timer()
'l1-------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l1.Y2 < 6735 Then
l1.Visible = True
l1.Y2 = l1.Y2 + 120
l1.X2 = l1.X2 + 120
End If
'l1--------------------------------

End Sub



Private Sub Timer10_Timer()
If Picture2.Top > -5280 Then
Picture2.Top = Picture2.Top - 40
cmd.Top = cmd.Top - 40
End If
End Sub



Private Sub Timer11_Timer()
Text2.Text = Time
End Sub


Private Sub Timer2_Timer()
'l2--------------------------------
Timer1.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False

i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l2.Y2 < 6735 Then
l2.Visible = True
l2.Y2 = l2.Y2 + 120
End If
'l2---------------------------------
End Sub

Private Sub Timer3_Timer()
'l3---------------------------------
Timer2.Enabled = False
Timer1.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l3.Y2 < 6735 Then
l3.Visible = True
l3.Y2 = l3.Y2 + 120
End If
'l3---------------------------------

End Sub

Private Sub Timer4_Timer()
'l4---------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer1.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l4.Y2 < 6735 Then
l4.Visible = True
l4.Y2 = l4.Y2 + 120
End If
'l4---------------------------------

End Sub

Private Sub Timer5_Timer()
'l5---------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer1.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l5.X1 > 0 Then
l5.Visible = True
l5.X1 = l5.X1 - 120
l5.Y1 = l5.Y1 + 120
End If
'l5---------------------------------

End Sub

Private Sub Timer6_Timer()
'l6---------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer1.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l6.X2 < 6735 Then
l6.Visible = True
l6.X2 = l6.X2 + 120
End If
'l6---------------------------------

End Sub

Private Sub Timer7_Timer()
'l7---------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer1.Enabled = False
Timer8.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l7.X2 < 6735 Then
l7.Visible = True
l7.X2 = l7.X2 + 120
End If
'l7---------------------------------

End Sub

Private Sub Timer8_Timer()
'l8---------------------------------
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer1.Enabled = False
i1.Enabled = False
i2.Enabled = False
i3.Enabled = False
i4.Enabled = False
i5.Enabled = False
i6.Enabled = False
i7.Enabled = False
i8.Enabled = False
i9.Enabled = False
If l8.X2 < 6735 Then
l8.Visible = True
l8.X2 = l8.X2 + 120
End If
'l8---------------------------------
End Sub

Private Sub Timer9_Timer()
If i1 <> blank.Picture Then
If i2 <> blank.Picture Then
If i3 <> blank.Picture Then
If i4 <> blank.Picture Then
If i5 <> blank.Picture Then
If i6 <> blank.Picture Then
If i7 <> blank.Picture Then
If i8 <> blank.Picture Then
If i9 <> blank.Picture Then
If l1.Visible = False Then
If l2.Visible = False Then
If l3.Visible = False Then
If l4.Visible = False Then
If l5.Visible = False Then
If l6.Visible = False Then
If l7.Visible = False Then
If l8.Visible = False Then
pname.Caption = "Game Draw"
Image1.Visible = False
Label1.Visible = True

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
If pname.Caption = "Player 1 turn" Then Image1.Picture = x.Picture
If pname.Caption = "Player 2 turn" Then Image1.Picture = z.Picture
If pname.Caption = "Game Complete Player 2 Wins" Then Image1.Visible = False
If pname.Caption = "Game Complete Player 1 Wins" Then Image1.Visible = False
If l1.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l2.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l3.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l4.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l5.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l6.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l7.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If

If l8.Visible = True Then
If Text1.Text = "x" Then
pname.Caption = "Game Complete Player 2 Wins"
Else
If Text1.Text = "0" Then
pname.Caption = "Game Complete Player 1 Wins"
End If
End If
End If


'l1---------------------------------
If i1.Picture = z.Picture Then
If i5.Picture = z.Picture Then
If i9.Picture = z.Picture Then
Timer1.Enabled = True
End If
End If
End If
If i1.Picture = x.Picture Then
If i5.Picture = x.Picture Then
If i9.Picture = x.Picture Then
Timer1.Enabled = True
End If
End If
End If
'l1---------------------------------

'l2---------------------------------
If i1.Picture = z.Picture Then
If i4.Picture = z.Picture Then
If i7.Picture = z.Picture Then
Timer2.Enabled = True
End If
End If
End If
If i1.Picture = x.Picture Then
If i4.Picture = x.Picture Then
If i7.Picture = x.Picture Then
Timer2.Enabled = True
End If
End If
End If
'l2---------------------------------

'l3---------------------------------
If i2.Picture = z.Picture Then
If i5.Picture = z.Picture Then
If i8.Picture = z.Picture Then
Timer3.Enabled = True
End If
End If
End If
If i2.Picture = x.Picture Then
If i5.Picture = x.Picture Then
If i8.Picture = x.Picture Then
Timer3.Enabled = True
End If
End If
End If
'l3---------------------------------

'l4---------------------------------
If i3.Picture = z.Picture Then
If i6.Picture = z.Picture Then
If i9.Picture = z.Picture Then
Timer4.Enabled = True
End If
End If
End If
If i3.Picture = x.Picture Then
If i6.Picture = x.Picture Then
If i9.Picture = x.Picture Then
Timer4.Enabled = True
End If
End If
End If
'l4---------------------------------

'l5---------------------------------
If i3.Picture = z.Picture Then
If i5.Picture = z.Picture Then
If i7.Picture = z.Picture Then
Timer5.Enabled = True
End If
End If
End If
If i3.Picture = x.Picture Then
If i5.Picture = x.Picture Then
If i7.Picture = x.Picture Then
Timer5.Enabled = True
End If
End If
End If
'l5---------------------------------

'l6---------------------------------
If i1.Picture = z.Picture Then
If i2.Picture = z.Picture Then
If i3.Picture = z.Picture Then
Timer6.Enabled = True
End If
End If
End If
If i1.Picture = x.Picture Then
If i2.Picture = x.Picture Then
If i3.Picture = x.Picture Then
Timer6.Enabled = True
End If
End If
End If
'l6---------------------------------

'l7---------------------------------
If i4.Picture = z.Picture Then
If i5.Picture = z.Picture Then
If i6.Picture = z.Picture Then
Timer7.Enabled = True
End If
End If
End If
If i4.Picture = x.Picture Then
If i5.Picture = x.Picture Then
If i6.Picture = x.Picture Then
Timer7.Enabled = True
End If
End If
End If
'l7---------------------------------

'l8---------------------------------
If i7.Picture = z.Picture Then
If i8.Picture = z.Picture Then
If i9.Picture = z.Picture Then
Timer8.Enabled = True
End If
End If
End If
If i7.Picture = x.Picture Then
If i8.Picture = x.Picture Then
If i9.Picture = x.Picture Then
Timer8.Enabled = True
End If
End If
End If
'l8---------------------------------
If l1.Visible = True Or l2.Visible = True Or l3.Visible = True Or l4.Visible = True Or l5.Visible = True Or l6.Visible = True Or l7.Visible = True Or l8.Visible = True Then
Label1.Visible = True
End If
End Sub
