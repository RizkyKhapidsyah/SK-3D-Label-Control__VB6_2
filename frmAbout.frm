VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Label3DControl.Label3D Label3DAddress 
      Height          =   570
      Left            =   1725
      TabIndex        =   4
      ToolTipText     =   "Click here to send me mail."
      Top             =   2595
      Width           =   3255
      _ExtentX        =   2461
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "jmeller@home.com"
      Caption         =   "jmeller@home.com"
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99
   End
   Begin Label3DControl.Label3D Label3DEmail 
      Height          =   510
      Left            =   300
      TabIndex        =   3
      Top             =   2595
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "e-mail:"
      Caption         =   "e-mail:"
      Color1          =   -2147483632
      AutoSize        =   0
   End
   Begin Label3DControl.Label3D Label3DAuthor 
      Height          =   600
      Left            =   68
      TabIndex        =   2
      Top             =   1455
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Written by John Meller"
      Caption         =   "Written by John Meller"
      Color1          =   -2147483632
      Color2          =   -2147483628
      AutoSize        =   0
   End
   Begin Label3DControl.Label3D Label3DTitle 
      Height          =   1350
      Left            =   480
      TabIndex        =   1
      Top             =   45
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   2381
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dayton"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color1          =   -2147483635
      Color2          =   -2147483628
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2010
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3DAddress.Color1 = vbRed

End Sub

Private Sub Label3DAddress_Click()
    Shell "Start.exe " & "mailto:jmeller@home.com?Subject=Hello", 0

End Sub

Private Sub Label3DAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3DAddress.Color1 = vbBlue
End Sub
