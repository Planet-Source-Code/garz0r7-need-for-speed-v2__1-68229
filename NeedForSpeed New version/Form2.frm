VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::NeedForSpeed:: Don't Forget To Vote........!^&$*"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   120
   End
   Begin VB.PictureBox goodbye 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   0
      Picture         =   "Form2.frx":030A
      ScaleHeight     =   3315
      ScaleWidth      =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   4830
   End
   Begin MCI.MMControl MM 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox what 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   $"Form2.frx":3495
      Top             =   3600
      Width           =   6375
   End
   Begin VB.Label WHAT1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "You Can Vote It HERE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo er

goodbye.Top = WHAT1.Height + what.Height

'Form1.Hide
Form2.Show

'IF YOU HAVE PROBLEM
'THEN ERASE THESE 3 LINES!!!
'****************************
MM.FileName = App.Path + "/rot.wav"
MM.Command = "open"
MM.Command = "play"
Form2.Timer1.Interval = 100

er:

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Hide


End Sub

Private Sub Timer1_Timer()

If goodbye.Top >= 0 Then
goodbye.Top = goodbye.Top - 30
WHAT1.Top = goodbye.Top + goodbye.Height
WHAT1.Width = goodbye.Width
what.Top = WHAT1.Top + WHAT1.Height
what.Width = goodbye.Width
End If

End Sub
