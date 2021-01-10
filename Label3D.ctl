VERSION 5.00
Begin VB.UserControl Label3D 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   PropertyPages   =   "Label3D.ctx":0000
   ScaleHeight     =   1080
   ScaleWidth      =   1470
   ToolboxBitmap   =   "Label3D.ctx":0014
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Label"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Label"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   1290
   End
End
Attribute VB_Name = "Label3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MyBackStyle
    Transparent = 0
    Opaque = 1
End Enum

Public Enum MyBorderStyle
    None = 0
    [Fixed Single] = 1
End Enum

Public Enum MyAutoSize
    [True]
    [False]
End Enum

' I just want to have the custom pointer
' if you wish add to the list
Public Enum MyMousePointer
    None = 0
    [Custom] = 99
End Enum

'Default Property Values:
Const m_def_ShadowLeft = 25
Const m_def_ShadowTop = 25
Const m_def_AutoSize = 1

'Property Variables:
Dim m_ShadowLeft As Integer
Dim m_ShadowTop As Integer
Dim m_AutoSize As MyAutoSize

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows the About Dialog Box"
Attribute ShowAbout.VB_UserMemId = -552
    frmAbout.Show vbModal
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label(0),Label,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = Label(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label(0).Font = New_Font
    Set Label(1).Font = New_Font
    Call UserControl_Resize
    
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As MyBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As MyBackStyle)
'    On Error GoTo BackStyleError
    
    If New_BackStyle = Opaque Or New_BackStyle = Transparent Then
        UserControl.BackStyle() = New_BackStyle
        PropertyChanged "BackStyle"
    Else
        err.Raise Number:=vbObjectError + 1001, _
                     Description:="Invalid BackStyle value (0 or 1 Only)"
    End If
    
'BackStyleError:
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As MyBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MyBorderStyle)
    
    If New_BorderStyle = 0 Or New_BorderStyle = 1 Then
        UserControl.BorderStyle() = New_BorderStyle
        PropertyChanged "BorderStyle"
    Else
        err.Raise Number:=vbObjectError + 1002, _
                     Description:="Invalid BorderStyle value (0 or 1 Only)"
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


Private Sub Label_Click(Index As Integer)
    ' a little cheat here :)
    If Index = 0 Or Index = 1 Then Call UserControl_Click
    
End Sub

Private Sub Label_DblClick(Index As Integer)
    If Index = 0 Or Index = 1 Then Call UserControl_DblClick

End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label(0),Label,0,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
    Caption = Label(0).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label(0).Caption() = New_Caption
    Label(1).Caption() = New_Caption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Resize()
    If m_AutoSize = 0 Then Call SizeControl
    
    'center the labels
    Label(0).Left = (UserControl.ScaleWidth - (Label(0).Width + m_ShadowLeft)) / 2
    Label(0).Top = (UserControl.ScaleHeight - (Label(0).Height + m_ShadowTop)) / 2
    
    Label(1).Left = Label(0).Left + m_ShadowLeft
    Label(1).Top = Label(0).Top + m_ShadowTop
   
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label(0),Label,0,ForeColor
Public Property Get Color1() As OLE_COLOR
Attribute Color1.VB_Description = "Returns/sets the foreground color used in the top caption "
Attribute Color1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Color1 = Label(0).ForeColor
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    Label(0).ForeColor() = New_Color1
    PropertyChanged "Color1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label(1),Label,1,ForeColor
Public Property Get Color2() As OLE_COLOR
Attribute Color2.VB_Description = "Returns/sets the foreground color used to the bottom caption (shadow)"
Attribute Color2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Color2 = Label(1).ForeColor
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    Label(1).ForeColor() = New_Color2
    PropertyChanged "Color2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,25
Public Property Get ShadowLeft() As Integer
Attribute ShadowLeft.VB_Description = "The position of Label(1) left side. This is the 'Shadow'."
Attribute ShadowLeft.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowLeft = m_ShadowLeft
End Property

Public Property Let ShadowLeft(ByVal New_ShadowLeft As Integer)
    m_ShadowLeft = New_ShadowLeft
    Label(1).Left = Label(0).Left + New_ShadowLeft
    Call UserControl_Resize
    PropertyChanged "ShadowLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,25
Public Property Get ShadowTop() As Integer
Attribute ShadowTop.VB_Description = "The position of label(1) top side. This is the 'Shadow'."
Attribute ShadowTop.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowTop = m_ShadowTop
End Property

Public Property Let ShadowTop(ByVal New_ShadowTop As Integer)
    m_ShadowTop = New_ShadowTop
    Label(1).Top = Label(0).Top + New_ShadowTop
    Call UserControl_Resize
    PropertyChanged "ShadowTop"
End Property

Public Property Get AutoSize() As MyAutoSize
Attribute AutoSize.VB_Description = "Set the AutoSize property to True or False"
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal New_AutoSize As MyAutoSize)
    If New_AutoSize = 1 Or New_AutoSize = 0 Then
        m_AutoSize = New_AutoSize

        PropertyChanged "AutoSize"
        Call SizeControl
    Else
        Debug.Print m_AutoSize
        err.Raise Number:=vbObjectError + 1003, _
                     Description:="Invalid AutoSize value (0 or 1 Only)"
    End If
        Debug.Print m_AutoSize
End Property
' I could have just as easily set this up so that the usercontrol mouseicon and pointer were set first
' but i mapped the label first...

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    Set Label(0).MouseIcon = UserControl.MouseIcon
    Set Label(1).MouseIcon = UserControl.MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MyMousePointer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MyMousePointer)
    UserControl.MousePointer() = New_MousePointer
    Label(0).MousePointer() = UserControl.MousePointer
    Label(1).MousePointer() = UserControl.MousePointer
    PropertyChanged "MousePointer"
End Property



Private Sub SizeControl()
    'if the shadow is negative (on the left side or top side)
    'then we change it a positive number
    
     If m_ShadowLeft < 0 Then
        UserControl.Width = Label(0).Width + (m_ShadowLeft * -1)
    Else
        UserControl.Width = Label(0).Width + m_ShadowLeft
    End If
    
    If m_ShadowTop < 0 Then
        UserControl.Height = Label(0).Height + (m_ShadowTop * -1)
    Else
        UserControl.Height = Label(0).Height + m_ShadowTop
    End If
   
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ShadowLeft = m_def_ShadowLeft
    m_ShadowTop = m_def_ShadowTop
    m_AutoSize = m_def_AutoSize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Label(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Label(1).Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label(0).Caption = PropBag.ReadProperty("Caption", "3D Label")
    Label(1).Caption = PropBag.ReadProperty("Caption", "3D Label")
    Label(0).ForeColor = PropBag.ReadProperty("Color1", &HFF&)
    Label(1).ForeColor = PropBag.ReadProperty("Color2", &H0&)
    m_ShadowLeft = PropBag.ReadProperty("ShadowLeft", m_def_ShadowLeft)
    m_ShadowTop = PropBag.ReadProperty("ShadowTop", m_def_ShadowTop)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Label(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", Label(1).Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", Label(0).Caption, "3D Label")
    Call PropBag.WriteProperty("Caption", Label(1).Caption, "3D Label")
    Call PropBag.WriteProperty("Color1", Label(0).ForeColor, &HFF&)
    Call PropBag.WriteProperty("Color2", Label(1).ForeColor, &H0&)
    Call PropBag.WriteProperty("ShadowLeft", m_ShadowLeft, m_def_ShadowLeft)
    Call PropBag.WriteProperty("ShadowTop", m_ShadowTop, m_def_ShadowTop)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)

  
End Sub

