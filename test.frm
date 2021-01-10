VERSION 5.00
Object = "*\ALabel3DControl.vbp"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin Label3DControl.Label3D Label3D1 
      Height          =   2115
      Left            =   1365
      TabIndex        =   0
      Top             =   990
      Width           =   6030
      _extentx        =   10636
      _extenty        =   3731
      font            =   "test.frx":0000
      font            =   "test.frx":002C
      backstyle       =   1
      shadowleft      =   105
      shadowtop       =   95
      autosize        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3D1.Color1 = vbRed

End Sub

Private Sub Label3D1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3D1.Color1 = vbBlue
    
End Sub
