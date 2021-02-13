VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0031000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6888
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8592
   LinkTopic       =   "Form1"
   ScaleHeight     =   6888
   ScaleWidth      =   8592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H0031000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      ItemData        =   "Form1.frx":0000
      Left            =   3360
      List            =   "Form1.frx":001D
      TabIndex        =   7
      Text            =   "5"
      Top             =   3360
      Width           =   1332
   End
   Begin MSCommLib.MSComm port 
      Left            =   6240
      Top             =   2760
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin vb6Arduino.Button Button1 
      Height          =   732
      Left            =   1680
      TabIndex        =   0
      Top             =   5280
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   1291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "LED ON"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":003B
      Picture         =   "Form1.frx":0057
      FixedSingle     =   0
   End
   Begin vb6Arduino.Button Button2 
      Height          =   732
      Left            =   4560
      TabIndex        =   1
      Top             =   5280
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   1291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "LED OFF"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":0073
      Picture         =   "Form1.frx":008F
      FixedSingle     =   0
   End
   Begin vb6Arduino.Button Button3 
      Height          =   732
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   4932
      _ExtentX        =   8700
      _ExtentY        =   1291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Connect to Port"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":00AB
      Picture         =   "Form1.frx":00C7
      FixedSingle     =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Developed By Sihab Sahariar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   8652
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Port :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   492
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   1692
   End
   Begin VB.Label led 
      BackStyle       =   0  'Transparent
      Caption         =   "LED IS OFF"
      Height          =   252
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   852
   End
   Begin VB.Image Image1 
      Height          =   2412
      Left            =   2880
      Picture         =   "Form1.frx":00E3
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "VB6 with Arduino"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Integer
Dim yy As Integer
Dim LEDon As String, LEDoff As String


Private Sub Button1_Click()
LEDon = "LEDON"
port.Output = LEDon
led.Caption = "LED IS ON"
End Sub

Private Sub Button2_Click()
LEDoff = "LEDOFF"
port.Output = LEDoff
led.Caption = "LED IS OFF"
End Sub

Private Sub Button3_Click()
On Error GoTo s
port.PortOpen = True
MsgBox "Port Opened", vbInformation, "Successul"
s:
    MsgBox "Please check if the port is open or not", vbCritical, "Error"
    

End Sub

Private Sub Combo1_Click()
port.CommPort = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Form_Load()
port.CommPort = 5 'Add you arduino port, You can find from device manager.
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Me.Left = Me.Left - xx + X
Me.Top = Me.Top - yy + Y
End If
End Sub

Private Sub Label2_Click()
End
End Sub

