VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   ":: Need For Speed :: *Parazan Productions*"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   7800
      Top             =   1560
   End
   Begin VB.PictureBox vgear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7680
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   510
      ScaleWidth      =   525
      TabIndex        =   39
      Top             =   240
      Width           =   555
   End
   Begin VB.TextBox ai 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   37
      Text            =   "96"
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CheckBox second 
      BackColor       =   &H80000007&
      Caption         =   "Race With Another Car?"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4080
      TabIndex        =   35
      Top             =   5640
      Width           =   2055
   End
   Begin VB.PictureBox car2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   1440
      Picture         =   "Form1.frx":07FD
      ScaleHeight     =   795
      ScaleWidth      =   360
      TabIndex        =   33
      Top             =   5040
      Width           =   390
   End
   Begin VB.PictureBox car 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   480
      Picture         =   "Form1.frx":0C94
      ScaleHeight     =   795
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   390
   End
   Begin VB.PictureBox started 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   120
      Picture         =   "Form1.frx":112B
      ScaleHeight     =   30
      ScaleWidth      =   2250
      TabIndex        =   32
      Top             =   5040
      Width           =   2280
   End
   Begin VB.CheckBox auto 
      BackColor       =   &H80000007&
      Caption         =   "Automatic Gear?"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2520
      MaskColor       =   &H00808080&
      TabIndex        =   30
      Top             =   5640
      Width           =   3615
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   960
      Picture         =   "Form1.frx":14EA
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   27
      Top             =   6360
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   1680
      Picture         =   "Form1.frx":18BB
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   26
      Top             =   6480
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   1680
      Picture         =   "Form1.frx":1C7B
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   25
      Top             =   7080
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   360
      Picture         =   "Form1.frx":2040
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   24
      Top             =   7080
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   1200
      Picture         =   "Form1.frx":2406
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   23
      Top             =   6720
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":27C4
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   22
      Top             =   6720
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   960
      Picture         =   "Form1.frx":2B86
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   21
      Top             =   6960
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   480
      Picture         =   "Form1.frx":2F3F
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   20
      Top             =   6360
      Width           =   840
   End
   Begin VB.PictureBox konos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   2040
      Picture         =   "Form1.frx":3303
      ScaleHeight     =   195
      ScaleWidth      =   810
      TabIndex        =   19
      Top             =   6720
      Width           =   840
   End
   Begin VB.PictureBox finish 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      Picture         =   "Form1.frx":36C5
      ScaleHeight     =   60
      ScaleWidth      =   2250
      TabIndex        =   18
      Top             =   6240
      Width           =   2280
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start New Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Timer Timer3 
      Left            =   7680
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Left            =   8160
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Left            =   7200
      Top             =   3360
   End
   Begin VB.PictureBox lines 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   6150
      Index           =   1
      Left            =   960
      Picture         =   "Form1.frx":3A84
      ScaleHeight     =   6120
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox lines 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   6150
      Index           =   0
      Left            =   960
      Picture         =   "Form1.frx":3D83
      ScaleHeight     =   6120
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox road 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   6150
      Left            =   120
      Picture         =   "Form1.frx":4082
      ScaleHeight     =   6120
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2160
   End
   Begin VB.PictureBox back 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   6150
      Left            =   0
      Picture         =   "Form1.frx":46AD
      ScaleHeight     =   6120
      ScaleWidth      =   2190
      TabIndex        =   31
      Top             =   0
      Width           =   2220
   End
   Begin VB.Shape tgear 
      BorderColor     =   &H00000000&
      Height          =   855
      Left            =   7320
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape l1 
      BorderColor     =   &H000000C0&
      Height          =   855
      Left            =   7680
      Top             =   480
      Width           =   15
   End
   Begin VB.Shape m1 
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   7920
      Top             =   480
      Width           =   15
   End
   Begin VB.Shape r1 
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   8280
      Top             =   480
      Width           =   15
   End
   Begin VB.Shape mm 
      BorderColor     =   &H000000C0&
      Height          =   15
      Left            =   7320
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      Height          =   1815
      Left            =   6360
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Artificial Intelligence ( AI ) % (Must be 90-100):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   6600
      TabIndex        =   38
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label jtime 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label td 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   34
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      Height          =   1815
      Left            =   2400
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1455
      Left            =   2400
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   1935
      Left            =   2400
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   5160
      X2              =   4680
      Y1              =   1080
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   3240
      X2              =   2760
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Shape table1 
      Height          =   1095
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape table2 
      Height          =   1215
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label best 
      Height          =   495
      Left            =   9480
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label rpm2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form1.frx":50A3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   6600
      TabIndex        =   17
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label st 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label time0100 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label timekm 
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label a 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label g 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label te 
      Height          =   375
      Left            =   9480
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label real2 
      Caption         =   "0"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label real1 
      Caption         =   "0"
      Height          =   255
      Left            =   9480
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label x 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label tim 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label z 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label u 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   2640
      Picture         =   "Form1.frx":513E
      Top             =   240
      Width           =   3330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================
'UPDATES
'31-3-07
'
'1-GEARBOX DISPLAY.
'2-OPpONENT'S CAR HAS HIS OWN GEARBOX.
'3-car length correction.
'4-two different car selection(in folder 'images)
'==========================

'Th4nx f0r u51ng!

Option Explicit
Rem new
Dim speed2 As Double
Dim gear2 As Double
Dim distance2 As Double
Dim AI_difficulty As Double
Rem new

'your car gearbox
Dim geara(1 To 5, 1 To 3) As Double
Dim gearu(1 To 5, 1 To 3) As Double
'opponent's car gerbox
Dim geara2(1 To 5, 1 To 3) As Double
Dim gearu2(1 To 5, 1 To 3) As Double

Dim gear As Integer
Dim sp As Integer




Dim rpm As Long

Const car_meters = 4.5 'meters
Dim distance As Double 'km
Dim started_time As Double
Dim dt As Double
Dim ts2 As Double
Dim dt2 As Double
Dim ts As Double
Dim j As Integer
Dim steering As Double
'Const steering_max = 12
Dim speed As Double 'km/h

Dim speed_lose As Double 'when steering:m/s^2
Dim car_accelaration As Double 'm/s^2
Dim car_decelaration As Double 'm/s^2
Dim steering_accelaration As Double 'm/s^2
Dim steering_deceleration As Double  'm/s^2



Private Sub ai_Change()



AI_difficulty = Int(ai) / 100

End Sub

Private Sub Command1_Click()
Dim t11 As Double
build_interface
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0

t11 = Timer
Do
DoEvents
'wait 3 seconds
st = 3 - Int(Timer - t11)
Loop Until Timer - t11 >= 3
st = "GooOOO!!"

build_interface
start
End Sub

Private Sub Form_Load()

Form1.Show
Form2.Show

Randomize Timer
build_interface





End Sub
Sub build_interface()

build_gear 'read the your gearbox's information
build_gear2 'opponents

'set the road correctly
back.Left = 100
road.Left = back.Left + back.Width / 2 - road.Width / 2
road.Top = 0

'set the road lines correctly
lines(0).Top = road.Top
lines(0).Left = road.Left + road.Width / 2 - lines(0).Width / 2
lines(1).Top = road.Top - lines(0).Height
lines(1).Left = road.Left + road.Width / 2 - lines(1).Width / 2

'set the speedometer and rpm-meter correctly
Line1.X1 = table1.Left + table1.Width / 2
Line1.Y1 = table1.Top + table1.Height / 2
Line2.X1 = table2.Left + table2.Width / 2
Line2.Y1 = table2.Top + table2.Height / 2

'set the car correctly
car.Left = road.Left + road.Width / 2 - car.Width / 2
car.Top = road.Top + road.Height - 1.5 * car.Height

car2.Left = -2 * car.Width
car2.Top = road.Top + road.Height - 1.5 * car.Height

'if you play with another player,show his car!
If second.Value = vbChecked Then

car.Left = road.Left + road.Width / 4 - car.Width / 2
car.Top = road.Top + road.Height - 1.5 * car.Height

car2.Left = road.Left + road.Width / 2 + road.Width / 4 - car.Width / 2
car2.Top = car.Top

End If


'set the finish line
finish.Top = road.Top - finish.Height

'set the start-line
started.Top = car.Top - started.Height
started.Left = road.Left
started.Width = road.Width

'erase previous records
timekm = "" 'previous race time.
time0100 = "" 'time in which car reaches 100km/h
td = ""
jtime = ""

'fix speedometer and rpm-meter value
set_speedometer 0, 1 'fix speedometer value
set_rpm_meter 0 'fix rpm-meter value

'build gearbox display
m1.Top = tgear.Top
l1.Top = tgear.Top
r1.Top = tgear.Top

l1.Left = tgear.Left + tgear.Width / 4
m1.Left = tgear.Left + tgear.Width / 2
r1.Left = tgear.Left + tgear.Width * 3 / 4

l1.Height = tgear.Height
m1.Height = tgear.Height
r1.Height = tgear.Height / 2

mm.Left = l1.Left
mm.Top = tgear.Top + tgear.Height / 2
mm.Width = tgear.Width / 2

vgear.Top = tgear.Top
vgear.Left = tgear.Left
End Sub

Sub start()
Dim s As Integer
Randomize Timer
For s = 1 To 9
'konos(s).Left = road.Left + Rnd() * (road.Width - konos(s).Width)
konos(s).Left = road.Left
konos(s).Top = road.Top - konos(s).Height
Next

sp = 0




distance = 0 'distance that the car travelled
speed = 0 'initial speed of the car
steering = 0 'steering acceleration of the car
gear = 1 'set the first gear of the car

set_speedometer 0, 1 'fix speedometer value
car_accelaration = 0 'm/s^2.initial car acceleration.this doesn't really effect!

'car constants=============================
speed_lose = 0.5 'm/s^2.When braking,this will be the decelleration of the car
car_decelaration = 11 'm/s^2 car's decceleration when braking
steering_accelaration = 15 'm/s^2.car's steering acceleretion(how speedy our car is streering)
steering_deceleration = 17 'm/s^2.how fast the car recovers after steering.
'car constants ended========================

'setup the second car
speed2 = 0
gear2 = 1
distance2 = 0

'setup AI_difficalty
AI_difficulty = Int(ai) / 100

'initialize the timers
started_time = GetTickCount
ts = GetTickCount
ts2 = GetTickCount

'start the timers
Timer1.Interval = 10
Timer2.Interval = 10
Timer3.Interval = 200


End Sub
Sub build_gear()

'this is your gearbox info!!!

'the accelarations for eatch gearbox
'when the speed=0!!
geara(1, 1) = 3
geara(2, 1) = 0.5 * geara(1, 1)
geara(3, 1) = 0.33 * geara(1, 1)
geara(4, 1) = 0.17 * geara(1, 1)
geara(5, 1) = 0.1 * geara(1, 1)

'these values mast be bigger than these above!
'these are the max accelarartions in the speed below(part2)
'geara(1, 2) = 4.3
'geara(2, 2) = 3.4
'geara(3, 2) = 2.4
'geara(4, 2) = 1.7
'geara(5, 2) = 1.1

geara(1, 2) = 6
geara(2, 2) = 0.6 * geara(1, 2)
geara(3, 2) = 0.4 * geara(1, 2)
geara(4, 2) = 0.3 * geara(1, 2)
geara(5, 2) = 0.22 * geara(1, 2)

geara(1, 3) = 0
geara(2, 3) = 0
geara(3, 3) = 0
geara(4, 3) = 0
geara(5, 3) = 0

Rem===================(part2)

gearu(1, 1) = 0
gearu(2, 1) = 0
gearu(3, 1) = 0
gearu(4, 1) = 0
gearu(5, 1) = 0

'here are the velocity of each gear where you have the
'biggest accelaration!So in this speed you must
'shift up the gear!
gearu(1, 2) = 50
gearu(2, 2) = 80
gearu(3, 2) = 110
gearu(4, 2) = 140
gearu(5, 2) = 170
'these values mast be bigger than these above!
'these are the biggest velocities for eatch gear!!
gearu(1, 3) = 70
gearu(2, 3) = 110
gearu(3, 3) = 155
gearu(4, 3) = 185
gearu(5, 3) = 220

End Sub
Sub build_gear2()

'this is opponent's car gearbox info!!!

'the accelarations for eatch gearbox
'when the speed=0!!
geara2(1, 1) = 3.3
geara2(2, 1) = 0.5 * geara2(1, 1)
geara2(3, 1) = 0.33 * geara2(1, 1)
geara2(4, 1) = 0.17 * geara2(1, 1)
geara2(5, 1) = 0.1 * geara2(1, 1)

'these values mast be bigger than these above!
'these are the max accelarartions in the speed below(part2)
'geara(1, 2) = 4.3
'geara(2, 2) = 3.4
'geara(3, 2) = 2.4
'geara(4, 2) = 1.7
'geara(5, 2) = 1.1

geara2(1, 2) = 6
geara2(2, 2) = 0.6 * geara2(1, 2)
geara2(3, 2) = 0.4 * geara2(1, 2)
geara2(4, 2) = 0.3 * geara2(1, 2)
geara2(5, 2) = 0.22 * geara2(1, 2)

geara2(1, 3) = 0
geara2(2, 3) = 0
geara2(3, 3) = 0
geara2(4, 3) = 0
geara2(5, 3) = 0

Rem===================(part2)

gearu2(1, 1) = 0
gearu2(2, 1) = 0
gearu2(3, 1) = 0
gearu2(4, 1) = 0
gearu2(5, 1) = 0

'here are the velocity of each gear where you have the
'biggest accelaration!So in this speed you must
'shift up the gear!
gearu2(1, 2) = 50
gearu2(2, 2) = 90
gearu2(3, 2) = 110
gearu2(4, 2) = 140
gearu2(5, 2) = 170
'these values mast be bigger than these above!
'these are the biggest velocities for eatch gear!!
gearu2(1, 3) = 70
gearu2(2, 3) = 110
gearu2(3, 3) = 155
gearu2(4, 3) = 185
gearu2(5, 3) = 220

End Sub

Function lamda_down(ByVal k As Integer) As Double
lamda_down = -geara(k, 2) / (gearu(k, 3) - gearu(k, 2))
End Function
Function lamda_up(ByVal k As Integer) As Double
lamda_up = (geara(k, 2) - geara(k, 1)) / gearu(k, 2)
End Function
Function best_gear(ByVal g As Integer) As Double
'returns the speed in which we have to
'shift up the gearbox,in order to have
'the best acceleration.

If g < 5 And g >= 1 Then
best_gear = (geara(g + 1, 1) - geara(g, 2) + lamda_down(g) * gearu(g, 2)) / (lamda_down(g) - lamda_up(g + 1))
End If

If g = 5 Then
'this will never been reached
best_gear = 1.2 * gearu(5, 3)
End If

If g = 0 Then
best_gear = 0
End If

End Function
Function lamda_down2(ByVal k As Integer) As Double
lamda_down2 = -geara2(k, 2) / (gearu2(k, 3) - gearu2(k, 2))
End Function
Function lamda_up2(ByVal k As Integer) As Double
lamda_up2 = (geara2(k, 2) - geara2(k, 1)) / gearu2(k, 2)
End Function
Function best_gear2(ByVal g As Integer) As Double
'returns the speed in which we have to
'shift up the gearbox,in order to have
'the best acceleration.

If g < 5 And g >= 1 Then
best_gear2 = (geara2(g + 1, 1) - geara2(g, 2) + lamda_down2(g) * gearu2(g, 2)) / (lamda_down2(g) - lamda_up2(g + 1))
End If

If g = 5 Then
'this will never been reached
best_gear2 = 1.2 * gearu2(5, 3)
End If

If g = 0 Then
best_gear2 = 0
End If

End Function




Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Sub automatic_gear()
'this will change the gearbox with the best way!

If speed >= best_gear(gear) And gear < 5 Then
gear = gear + 1
End If

If gear > 1 And speed < best_gear(gear - 1) Then
gear = gear - 1
End If

End Sub





Private Sub Timer1_Timer()

dt = GetTickCount - ts
ts = GetTickCount
dt = dt / 1000


best = Round(best_gear(gear), 2)
best = "Best Speed To Shift Up The Gear: " + best + " Km/h"

If auto.Value = vbChecked Then
'Let the computer change the gear.
automatic_gear
End If


'u = Round(((car_meters * speed / car.Height) * 3600 / 1000), 2)
'u = u + " Km/h"
u = Round(speed, 2)
u = "Speed:" + u + " Km/h"

'this steering the car to the desired direction
car.Left = car.Left + (steering * car.Height * 1000 / (3600 * car_meters)) * dt

If car.Left < road.Left Then
car.Left = road.Left
speed = speed - 20 * dt
End If

If car.Left + car.Width > road.Left + road.Width Then
car.Left = road.Left + road.Width - car.Width
speed = speed - 20 * dt
End If

If speed < 0 Then
speed = 0
End If

'check if you have gearbox overrun!!
check_gear dt



For j = 0 To 1
'this is moving the road lines
lines(j).Top = lines(j).Top + (speed * car.Height * 1000 / (3600 * car_meters)) * dt

If lines(j).Top > road.Height + road.Top Then
lines(j).Top = road.Top - lines(j).Height
End If

If lines(j).Top > lines(Abs(j - 1)).Top Then
lines(Abs(j - 1)).Top = lines(j).Top - lines(j).Height - 50
End If

Next

If second.Value = vbChecked Then
'if you play with another player
'then move him!
car2.Left = road.Left + road.Width / 2 + road.Width / 4 - car.Width / 2
move_second_car dt
End If




If started.Top <= road.Top + road.Height Then
'move the start-line..
started.Top = started.Top + (speed * car.Height * 1000 / (3600 * car_meters)) * dt
End If

If (1 - distance) * car.Height * 1000 / car_meters <= car.Top - road.Top Then
'move the finish-line..
finish.Top = finish.Top + (speed * car.Height * 1000 / (3600 * car_meters)) * dt
End If

'update car's distance so far..
distance = distance + speed * dt / 3600
x = Round(distance, 3)
x = "Distance: " + x + "Km"

'update racing time so far
tim = Round((GetTickCount - started_time) / 1000, 3)
tim = "Time: " + tim + " secs"

If distance >= 1 Then
'race completed::1km was runned!!!
timekm = Round((GetTickCount - started_time) / 1000, 3)
timekm = "Race Completed!Time:" + timekm + " secs"
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
End If

If speed >= 100 And sp = 0 Then
'reached 100km/h!!!
sp = 1
time0100 = Round((GetTickCount - started_time) / 1000, 3)
time0100 = "0-100 km/h in : " + time0100 + " secs"
End If


For j = 1 To 9
'MOVE OBJECTS.e.t 100m label,200m label.....
If ((10 - j) * 0.1 - distance) * car.Height * 1000 / car_meters <= car.Top - road.Top Then
konos(j).Top = konos(j).Top + (speed * car.Height * 1000 / (3600 * car_meters)) * dt
End If

If konos(j).Top <= car.Top And konos(j).Top + konos(j).Height >= car.Top Then
'SAVE STAGE'S TIME
jtime = "Last Stage's Time: " + tim
End If

Next


End Sub
Sub set_speedometer(ByVal speed As Double, ByVal gear As Integer)

Const min_angle = 3.14 / 1.25
Const max_angle = 250

Rem table1
Line1.X2 = Cos(min_angle + 2 * 3.14 * speed * max_angle / (gearu(gear, 3) * 360)) * table1.Width / 2 + table1.Left + table1.Width / 2
Line1.Y2 = Sin(min_angle + 2 * 3.14 * speed * max_angle / (gearu(gear, 3) * 360)) * table1.Height / 2 + table1.Top + table1.Height / 2
End Sub

Sub set_rpm_meter(ByVal speed As Double)

Const min_angle = 3.14 / 1.25
Const max_angle = 250

Rem table2
Line2.X2 = Cos(min_angle + 2 * 3.14 * speed * max_angle / (gearu(5, 3) * 360)) * table2.Width / 2 + table2.Left + table2.Width / 2
Line2.Y2 = Sin(min_angle + 2 * 3.14 * speed * max_angle / (gearu(5, 3) * 360)) * table2.Height / 2 + table2.Top + table2.Height / 2
End Sub

Sub check_gear(ByVal dt As Double)
Dim k As Integer
Dim temp As Integer
Dim lamda As Double
Dim angle As Double

temp = 0
For k = 1 To 3
If speed >= gearu(gear, k) Then
temp = k
End If
Next

'te = temp
g = gear
g = "Gear:" + g
If temp = 3 Then
car_accelaration = -3 'm/s
speed = speed + dt * car_accelaration * 3600 / 1000
Else
lamda = (gearu(gear, temp + 1) - gearu(gear, temp)) / (geara(gear, temp + 1) - geara(gear, temp))
car_accelaration = (speed - gearu(gear, temp)) / lamda + geara(gear, temp)
End If

'FIX SPEEDOMETER'S VALUE
set_speedometer speed, gear


'FIX RPM-METER VALUE
set_rpm_meter speed

rpm = speed * 8000 / gearu(gear, 3)

rpm2 = rpm
rpm2 = "Revolutions Per Minute: " + rpm2 + " Rpm"
a = Round(car_accelaration, 2)
a = "Car Accelaration:" + a + " m/s^2"
End Sub
Private Sub Timer2_Timer()


dt2 = GetTickCount - ts2
ts2 = GetTickCount
dt2 = dt2 / 1000

If GetAsyncKeyState(vbKeyLeft) Then
steering = steering - steering_accelaration * dt2 * 3600 / 1000
speed = speed - speed_lose * dt2 * 3600 / 1000
Else
If steering < 0 Then
steering = steering + steering_deceleration * dt2
End If
End If

If GetAsyncKeyState(vbKeyRight) Then
steering = steering + steering_accelaration * dt2 * 3600 / 1000
speed = speed - speed_lose * dt2 * 3600 / 1000
Else
If steering > 0 Then
steering = steering - steering_deceleration * dt2
End If
End If

If GetAsyncKeyState(vbKeyUp) Then
speed = speed + car_accelaration * dt2 * 3600 / 1000
Else
car_accelaration = -0.4
speed = speed + car_accelaration * dt2 * 3600 / 1000
End If

If GetAsyncKeyState(vbKeyDown) Then
speed = speed - car_decelaration * dt2 * 3600 / 1000
End If



If steering > steering_accelaration Then
steering = steering_accelaration
End If

If steering < -steering_accelaration Then
steering = -steering_accelaration
End If




z = Round(steering, 2)
z = "Streering Accelararion: " + z + " m/s^2"

'ts2 = GetTickCount
End Sub

Private Sub Timer3_Timer()
'this changes the gear!

If GetAsyncKeyState(vbKeyA) And gear < 5 Then
gear = gear + 1
End If

If GetAsyncKeyState(vbKeyZ) And gear > 1 Then
gear = gear - 1
End If

End Sub

Private Sub move_second_car(ByVal dt As Double)
Dim k As Integer
Dim temp As Integer
Dim lamda As Double
Dim accelaration As Double



'auto-gear-system for the second car..

If speed2 >= best_gear2(gear2) And gear2 < 5 Then
gear2 = gear2 + 1
End If

If gear2 > 1 And speed2 < best_gear2(gear2 - 1) Then
gear2 = gear2 - 1
End If

temp = 0
For k = 1 To 3
If speed2 >= gearu2(gear2, k) Then
temp = k
End If
Next

lamda = (gearu2(gear2, temp + 1) - gearu2(gear2, temp)) / (geara2(gear2, temp + 1) - geara2(gear2, temp))
accelaration = (speed2 - gearu2(gear2, temp)) / lamda + geara2(gear2, temp)

speed2 = speed2 + accelaration * AI_difficulty * dt * 3600 / 1000
distance2 = distance2 + speed2 * dt / 3600

'show the car!
car2.Top = car.Top - (Round(distance2 - distance, 5)) * 1000 * car.Height / car_meters

'avoid collision with the car1
If car.Left + car.Width + 50 > road.Left + road.Width / 2 Then
car.Left = road.Left + road.Width / 2 - car.Width - 50
End If



'time difference between the two cars,in seconds
td = Round(Round(distance2 - distance, 5) * 3600 / (speed + 0.00001), 3)

If second.Value = vbChecked Then
'if there is another player,show time difference
If td > 0 Then
td.ForeColor = vbRed
td = "You Are Second : + " + td + " secs"
Else
td.ForeColor = vbBlue
td = "You Are First : -" + td + " secs"
End If

End If

End Sub

Private Sub Timer4_Timer()


If gear = 1 Then
vgear.Top = tgear.Top
vgear.Left = tgear.Left
End If

If gear = 2 Then
vgear.Top = tgear.Top + tgear.Height - vgear.Height
vgear.Left = tgear.Left
End If

If gear = 3 Then
vgear.Top = tgear.Top
vgear.Left = tgear.Left + tgear.Width / 2 - vgear.Width / 2
End If

If gear = 4 Then
vgear.Top = tgear.Top + tgear.Height - vgear.Height
vgear.Left = tgear.Left + tgear.Width / 2 - vgear.Width / 2
End If

If gear = 5 Then
vgear.Top = tgear.Top
vgear.Left = tgear.Left + tgear.Width - vgear.Width
End If

End Sub


