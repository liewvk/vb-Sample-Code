VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Casino                     designed by Liew  Voon Kiong 2005"
   ClientHeight    =   7665
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "casino4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "casino4.frx":1272
   ScaleHeight     =   7665
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl2 
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Top             =   6360
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      PlayEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5880
      Top             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Press to Spin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VB Slot Machine"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"casino4.frx":24E4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   2
      Left            =   4080
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   1
      Left            =   2880
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   0
      Left            =   1680
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Key in amount to bet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu instruct 
         Caption         =   "Instructions"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim amount As Integer
Dim a, b, c As Integer





Private Sub Command1_Click()
Timer1.Enabled = True
MMControl1.Command = "Close"
MMControl2.Command = "close"

x = 0
Label2.Caption = "Your Credits"
amount = Val(Text1)
End Sub

Private Sub Command2_Click()
End
End Sub



Private Sub Form_Click()
Label3.Visible = False
End Sub

Private Sub Form_Load()
Label1.Caption = " Welcome to Play"

Label3.Visible = False

End Sub

Private Sub instruct_click()

Label3.Visible = True


End Sub

Private Sub Text1_Change()
amount = Val(Text1)

End Sub

Private Sub Timer1_Timer()
If x < 500 Then
spin
Else
Timer1.Enabled = False
MMControl1.Command = "Stop"
Label1.Alignment = 2

If (a = 3 And b = 3 And c <> 3) Or (a = 3 And c = 3 And b <> 3) Or (b = 3 And c = 3 And a <> 3) Then
Label1.Caption = " You win 20 dollars"
amount = amount + 20

End If

If (a = 4 And b = 4 And c <> 4) Or (a = 4 And c = 4 And b <> 4) Or (b = 4 And c = 4 And a <> 4) Then
Label1.Caption = " You win 30 dollars"
amount = amount + 30

End If

If (a = 5 And b = 5 And c <> 5) Or (a = 5 And c = 5 And b <> 5) Or (b = 5 And c = 5 And a <> 5) Then
Label1.Caption = " You win 40 dollars"
amount = amount + 40

End If

If (a = 3 And b = 3 And c = 3) Or (a = 4 And b = 4 And c = 4) Or (a = 5 And b = 5 And c = 5) Then

MMControl2.Notify = False
MMControl2.Wait = True
MMControl2.Shareable = False
MMControl2.DeviceType = "WaveAudio"
MMControl2.FileName = "D:\Liew Folder\VB program\audio\endgame.wav"
MMControl2.Command = "Open"
MMControl2.Command = "Play"


Label1.Caption = " Congratulation! Jackpot!!! You win 200 dollars!"
amount = amount + 200
End If

If (a = 3 And b = 4 And c = 5) Or (a = 3 And b = 5 And c = 4) Or (a = 4 And b = 3 And c = 5) Or (a = 4 And b = 5 And c = 3) Or (a = 5 And b = 4 And c = 3) Or (a = 5 And b = 3 And c = 4) Then

Label1.Caption = " Too bad, you lost 50 dollars"
amount = amount - 50
End If

If amount < 0 Then
Label1.Caption = "Oh! you're bankrupt!"
End If
Text1.Text = Str$(amount)
End If

End Sub
Sub spin()

x = x + 10
Randomize Timer
a = 3 + Int(Rnd * 3)
b = 3 + Int(Rnd * 3)
c = 3 + Int(Rnd * 3)

MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"
MMControl1.FileName = "D:\Liew Folder\VB program\audio\slot2.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"


Label1.Caption = "Good Luck!"
Label1.Alignment = a - 3
Shape1(0).Shape = a
If a = 3 Then
Shape1(0).FillColor = &HFF00&
End If
If a = 4 Then
Shape1(0).FillColor = &HFF00FF
End If
If a = 5 Then
Shape1(0).FillColor = &HFF0000

End If


Shape1(1).Shape = b
If b = 3 Then
Shape1(1).FillColor = &HFF00&
End If
If b = 4 Then
Shape1(1).FillColor = &HFF00FF
End If

If b = 5 Then
Shape1(1).FillColor = &HFF0000
End If

Shape1(2).Shape = c
If c = 3 Then
Shape1(2).FillColor = &HFF00&
End If
If c = 4 Then
Shape1(2).FillColor = &HFF00FF
End If
If c = 5 Then
Shape1(2).FillColor = &HFF0000
End If



End Sub
