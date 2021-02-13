VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   ScaleHeight     =   1500
   ScaleWidth      =   1920
   ToolboxBitmap   =   "Button1.ctx":0000
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Flag, Flag1 As Boolean
Private Normal_Images As StdPicture
Private Select_Images As StdPicture
Public Event Click()
Public Event MouseDown()
Public Event MouseMove()
Public Event MouseUp()

Property Get Enabled() As Boolean
    Enabled = Label3.Enabled
End Property

Property Let Enabled(ByVal Nval As Boolean)
    Label3.Enabled = Nval
    PropertyChanged "Enabled"
End Property

Property Get Text() As Variant
    Text = Label1.Caption
End Property
Property Let Text(ByVal Nval As Variant)
    Label1.Caption = Nval
    PropertyChanged "Text"
End Property

Property Get Font() As StdFont
    Set Font = Label1.Font
End Property
Property Set Font(ByVal Nval As StdFont)
    Set Label1.Font = Nval
    Set Label2.Font = Label1.Font
    PropertyChanged "Font"
End Property

Property Get Color() As OLE_COLOR
    Color = Label1.ForeColor
End Property
Property Let Color(ByVal Nval As OLE_COLOR)
    Label1.ForeColor = Nval
    PropertyChanged "Color"
End Property

Property Get Picture() As StdPicture
    Set Picture = Image1.Picture
    Set Normal_Images = Image1.Picture
End Property

Property Set Picture(ByVal Nval As StdPicture)
    Set Image1.Picture = Nval
    Set Normal_Images = Nval
    PropertyChanged "Picture"
End Property

Property Get Select_Image() As StdPicture
    Set Select_Image = Select_Images
End Property

Property Set Select_Image(ByVal Nval As StdPicture)
    Set Select_Images = Nval
    PropertyChanged "Select_Image"
End Property

Property Get FixedSingle() As Boolean
    FixedSingle = BorderStyle
End Property

Property Let FixedSingle(ByVal Nval As Boolean)
    If Nval = True Then
        BorderStyle = ccFixedSingle
    Else
        BorderStyle = ccNone
    End If
    PropertyChanged "FixedSingle"
End Property

Private Sub Label3_Click()
RaiseEvent Click
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Flag = False Then
    Flag = True
    If Me.FixedSingle = True Then
        Flag1 = True
    Else
        Flag1 = False
    End If
End If
If Flag1 = True Then
    Me.FixedSingle = False
Else
Me.FixedSingle = True
End If
Set Image1.Picture = Select_Images
RaiseEvent MouseDown
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Flag1 = True Then
    Me.FixedSingle = True
Else
Me.FixedSingle = False
End If
Set Image1.Picture = Normal_Images
RaiseEvent MouseUp
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    Label3.Top = 0
    Label3.Left = 0
    Label3.Width = Width
    Label3.Height = Height
    Label3.Visible = True
    
    Label2.AutoSize = True
    Set Label2.Font = Label1.Font
    
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Width
    Image1.Height = Height
    
    Label1.Width = Width
    Label1.Left = 0
    Label1.Height = (Height + Label2.Height) / 2
    Label1.Top = Height - Label1.Height
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Image1.Picture = PropBag.ReadProperty("Picture", Image1.Picture)
    Set Normal_Images = PropBag.ReadProperty("Picture", Normal_Images)
    Set Select_Images = PropBag.ReadProperty("Select_Image", Select_Image)
    Set Label1.Font = PropBag.ReadProperty("Font", Label1.Font)
    Label3.Enabled = PropBag.ReadProperty("Enabled", Label3.Enabled)
    Label1.Caption = PropBag.ReadProperty("Text", Label1.Caption)
    Label1.ForeColor = PropBag.ReadProperty("Color", Label1.ForeColor)
    BorderStyle = PropBag.ReadProperty("FixedSingle", BorderStyle)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Font", Label1.Font, Ambient.Font
    PropBag.WriteProperty "Text", Label1.Caption
    PropBag.WriteProperty "Enabled", Label3.Enabled
    PropBag.WriteProperty "Color", Label1.ForeColor
    PropBag.WriteProperty "Picture", Image1.Picture
    PropBag.WriteProperty "Picture", Normal_Images
    PropBag.WriteProperty "Select_Image", Select_Images
    PropBag.WriteProperty "FixedSingle", BorderStyle

End Sub
