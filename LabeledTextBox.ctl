VERSION 5.00
Begin VB.UserControl LabeledTextBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   ScaleHeight     =   240
   ScaleWidth      =   3690
   Begin VB.CommandButton cmdForm 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtForm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblAutoSizer 
      AutoSize        =   -1  'True
      Caption         =   "Ay"
      Height          =   195
      Left            =   3525
      TabIndex        =   3
      Top             =   135
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Line lineLabel 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   1470
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label lblForm 
      Caption         =   "Label"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "LabeledTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ButtonClick()
Event Change()
Event KeyPress(KeyAscii As Integer)

Private mbLookupButton As Boolean
Private mbLookupButtonOnFocus As Boolean
Private miLookupKey As Integer
Private miLookupKeyShift As Integer
Private mbLookupButtonInside As Boolean
Private mbUseContainerColor As Boolean
Private mnLockedBackColor As OLE_COLOR
Private mnLockedForeColor As OLE_COLOR
Private mnEntryBackColor As OLE_COLOR
Private mnEntryForeColor As OLE_COLOR
Private mnFocusedBackColor As OLE_COLOR
Private mnTextBoxLeft As Integer
Private mnProperHeight As Integer


Const OFFSET As Integer = 60

Public Enum LabelStyle
    ltsLabelLeft
    ltsLabelTop
    ltsLabelBottom
End Enum
Private meStyle As LabelStyle

Private Sub cmdForm_Click()
    RaiseEvent ButtonClick
End Sub

Private Sub txtForm_Change()
    RaiseEvent Change
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtForm_GotFocus
' DateTime  : 2003.Sep.21 05:00
' Author    : Adhimas Setianegara
' Purpose   : When the focus is on the textbox, highlight
'   the text
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Private Sub txtForm_GotFocus()
    txtForm.SelStart = 0
    txtForm.SelLength = Len(txtForm.Text) + 1
    
    If mbLookupButtonOnFocus Then ShowLookupButton
    If Not txtForm.Locked Then SetEntryBackColor mnFocusedBackColor
End Sub

Private Sub txtForm_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = miLookupKey) And (Shift = miLookupKeyShift) Then
        cmdForm.SetFocus
        cmdForm_Click
    End If
End Sub

Private Sub txtForm_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtForm_LostFocus()
    If mbLookupButtonOnFocus Then
        If UserControl.ActiveControl Is cmdForm Then RaiseEvent ButtonClick
        ShowLookupButton False
    End If
    
    SetEntryBackColor mnEntryBackColor
End Sub


'---------------------------------------------------------------------------------------
' Procedure : UserControl_AmbientChanged
' DateTime  : 2003.09.26 13:14
' Author    : Adhimas Setianegara
' Parameters:
'   none
' Purpose   :
'   Apply new color when the container property color changed
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_AmbientChanged(PropertyName As String)
    Select Case PropertyName
    Case "BackColor"
        If mbUseContainerColor Then lblForm.BackColor = _
            UserControl.Parent.BackColor
    
    Case "ForeColor"
        If mbUseContainerColor Then lblForm.ForeColor = _
            UserControl.Parent.ForeColor
        
    End Select
End Sub

Private Sub UserControl_InitProperties()
    lblForm.Caption = "Label:"
    lblForm.BackColor = UserControl.Parent.BackColor
    lblForm.ForeColor = UserControl.Parent.ForeColor
    mnEntryForeColor = txtForm.ForeColor
    SetEntryBackColor txtForm.BackColor
    mnLockedForeColor = mnEntryForeColor
    mnLockedBackColor = mnEntryBackColor
    mnFocusedBackColor = mnEntryBackColor
    mnTextBoxLeft = 1470
    
    miLookupKey = vbKeyF4
    miLookupKeyShift = 0
    meStyle = ltsLabelLeft
    
    mbUseContainerColor = True
    
    SetStyle
    ResizeTextBox
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Resize
' DateTime  : 2003.Sep.21 04:35
' Author    : Adhimas Setianegara
' Purpose   : Resize textbox portion when user resize
'   the user control.
' Parameters: none
'
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_Resize()
    If UserControl.Height <> mnProperHeight Then
        UserControl.Height = mnProperHeight
    Else
        ResizeTextBox
        ResizeLabel
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResizeTextBox
' DateTime  : 2003.Sep.21 05:11
' Author    : Adhimas Setianegara
' Purpose   : Resize textbox portion based on user control
'   width and the status of LookupButton
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Private Sub ResizeTextBox()
    If mbLookupButton Then
        cmdForm.Left = UserControl.Width - cmdForm.Width
        txtForm.Width = cmdForm.Left - txtForm.Left
    Else
        If UserControl.Width > txtForm.Left Then
            txtForm.Width = UserControl.Width - txtForm.Left
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Caption
' DateTime  : 2003.Sep.21 04:40
' Author    : Adhimas Setianegara
' Purpose   : Get the text label (lblForm)
'
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Get/Set the label for the textbox."

    Caption = lblForm.Caption

End Property

'---------------------------------------------------------------------------------------
' Procedure : Caption
' DateTime  : 2003.Sep.21 04:40
' Author    : Adhimas Setianegara
' Purpose   : change the text in label (lblForm)
'
' Parameters:
'   sLabel: the new label value
'---------------------------------------------------------------------------------------
'
Public Property Let Caption(ByVal asCaption As String)

    asCaption = Trim(asCaption)
    If Len(asCaption) > 0 Then
        If Right(asCaption, 1) <> ":" Then asCaption = asCaption + ":"
    End If
    
    lblForm.Caption = asCaption

    Call UserControl.PropertyChanged("Caption")

End Property

'---------------------------------------------------------------------------------------
' Procedure : Text
' DateTime  : 2003.Sep.21 04:40
' Author    : Adhimas Setianegara
' Purpose   : get the text value
'
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Public Property Get Text() As String

    Text = txtForm.Text

End Property

'---------------------------------------------------------------------------------------
' Procedure : Text
' DateTime  : 2003.Sep.21 04:40
' Author    : Adhimas Setianegara
' Purpose   : change the text value (txtForm)
'
' Parameters:
'   sText: the new text value
'---------------------------------------------------------------------------------------
'
Public Property Let Text(ByVal sText As String)

    txtForm.Text = sText

    Call UserControl.PropertyChanged("Text")

End Property

'---------------------------------------------------------------------------------------
' Procedure : LookupButton
' DateTime  : 2003.Sep.21 04:49
' Author    : Adhimas Setianegara
' Purpose   : Get the status of lookup button visibility
'
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Public Property Get LookupButton() As Boolean
Attribute LookupButton.VB_Description = "Show/hide lookup button."

    LookupButton = mbLookupButton

End Property

'---------------------------------------------------------------------------------------
' Procedure : LookupButton
' DateTime  : 2003.Sep.21 04:49
' Author    : Adhimas Setianegara
' Purpose   : Whether to show lookup button in the
'   right side of the control.
' Parameters:
'   abLookupButton: True/False
'---------------------------------------------------------------------------------------
'
Public Property Let LookupButton(ByVal abLookupButton As Boolean)

    If abLookupButton Then SetLocked False
    
    ShowLookupButton abLookupButton
    If abLookupButton Then mbLookupButtonOnFocus = False
    
    Call UserControl.PropertyChanged("LookupButton")
End Property

'---------------------------------------------------------------------------------------
' Procedure : MaxLength
' DateTime  : 2003.Sep.21 16:01
' Author    : Adhimas Setianegara
' Purpose   : Get max length of the value in the textbox
'
' Parameters:
'   none
'---------------------------------------------------------------------------------------
'
Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Maximum length of the text inside the textbox. Set to 0 to allow unlimited string length."

    MaxLength = txtForm.MaxLength

End Property

'---------------------------------------------------------------------------------------
' Procedure : MaxLength
' DateTime  : 2003.Sep.21 16:01
' Author    : Adhimas Setianegara
' Purpose   : Change max length of the value in the textbox
'
' Parameters:
'   aiMaxLength: The new maximum length
'---------------------------------------------------------------------------------------
'
Public Property Let MaxLength(ByVal aiMaxLength As Integer)

    txtForm.MaxLength = aiMaxLength

    Call UserControl.PropertyChanged("MaxLength")

End Property


'---------------------------------------------------------------------------------------
' Procedure : PasswordChar
' DateTime  : 2003.Sep.21 16:06
' Author    : Adhimas Setianegara
' Purpose   :
'
' Parameters:
'
'---------------------------------------------------------------------------------------
'
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Get/set the character that is used for password input. Leave this property empty for normal text input."

    PasswordChar = txtForm.PasswordChar

End Property

'---------------------------------------------------------------------------------------
' Procedure : PasswordChar
' DateTime  : 2003.Sep.21 16:06
' Author    : Adhimas Setianegara
' Purpose   :
'
' Parameters:
'
'---------------------------------------------------------------------------------------
'
Public Property Let PasswordChar(ByVal sPasswordChar As String)

    txtForm.PasswordChar = sPasswordChar

    Call UserControl.PropertyChanged("PasswordChar")

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblForm.Caption = PropBag.ReadProperty("Caption", "Label:")
    txtForm.Text = PropBag.ReadProperty("Text", "")
    txtForm.MaxLength = PropBag.ReadProperty("MaxLength", 10)
    txtForm.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    mbLookupButtonOnFocus = PropBag.ReadProperty("LookupButtonOnFocus", False)
    miLookupKey = PropBag.ReadProperty("LookupKey", vbKeyF4)
    miLookupKeyShift = PropBag.ReadProperty("LookupKeyShift", 0)
    mbLookupButtonInside = PropBag.ReadProperty("LookupButtonInside", False)
    mbUseContainerColor = PropBag.ReadProperty("UseContainerColor", True)
    lblForm.BackColor = PropBag.ReadProperty("BackColor", lblForm.BackColor)
    lblForm.ForeColor = PropBag.ReadProperty("ForeColor", lblForm.ForeColor)
    mnEntryBackColor = PropBag.ReadProperty("EntryBackColor", txtForm.BackColor)
    mnEntryForeColor = PropBag.ReadProperty("EntryForeColor", txtForm.ForeColor)
    mnLockedBackColor = PropBag.ReadProperty("LockedBackColor", txtForm.BackColor)
    mnLockedForeColor = PropBag.ReadProperty("LockedForeColor", txtForm.ForeColor)
    mnFocusedBackColor = PropBag.ReadProperty("FocusedBackColor", txtForm.BackColor)
    mnTextBoxLeft = PropBag.ReadProperty("TextBoxLeft", 1470)
    meStyle = PropBag.ReadProperty("Style", LabelStyle.ltsLabelLeft)
    
    SetFont PropBag.ReadProperty("Font", txtForm.Font)
    SetLocked PropBag.ReadProperty("Locked", False)
    SetStyle
    ShowLookupButton PropBag.ReadProperty("LookupButton", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "LookupButton", mbLookupButton
    PropBag.WriteProperty "Caption", lblForm.Caption, "Label"
    PropBag.WriteProperty "Text", txtForm.Text, ""
    PropBag.WriteProperty "MaxLength", txtForm.MaxLength, 10
    PropBag.WriteProperty "PasswordChar", txtForm.PasswordChar, ""
    PropBag.WriteProperty "Font", txtForm.Font
    PropBag.WriteProperty "LookupButtonOnFocus", mbLookupButtonOnFocus
    PropBag.WriteProperty "LookupKey", miLookupKey
    PropBag.WriteProperty "LookupKeyShift", miLookupKeyShift
    PropBag.WriteProperty "LookupButtonInside", mbLookupButtonInside
    PropBag.WriteProperty "UseContainerColor", mbUseContainerColor
    PropBag.WriteProperty "BackColor", lblForm.BackColor
    PropBag.WriteProperty "ForeColor", lblForm.ForeColor
    PropBag.WriteProperty "EntryBackColor", mnEntryBackColor
    PropBag.WriteProperty "EntryForeColor", mnEntryForeColor
    PropBag.WriteProperty "LockedBackColor", mnLockedBackColor
    PropBag.WriteProperty "LockedForeColor", mnLockedForeColor
    PropBag.WriteProperty "FocusedBackColor", mnFocusedBackColor
    PropBag.WriteProperty "Locked", txtForm.Locked
    PropBag.WriteProperty "TextBoxLeft", txtForm.Left
    PropBag.WriteProperty "Style", meStyle
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Font
' DateTime  : 2003.Sep.21 16:07
' Author    : Adhimas Setianegara
' Purpose   :
'
' Parameters:
'
'---------------------------------------------------------------------------------------
'
Public Property Get Font() As IFontDisp

    Set Font = txtForm.Font

End Property

'---------------------------------------------------------------------------------------
' Procedure : Font
' DateTime  : 2003.Sep.21 16:07
' Author    : Adhimas Setianegara
' Purpose   :
'
' Parameters:
'
'---------------------------------------------------------------------------------------
'
Public Property Set Font(objFont As IFontDisp)
    SetFont objFont
    
    Call UserControl.PropertyChanged("Font")
End Property


'---------------------------------------------------------------------------------------
' Procedure : LookupButtonOnFocus
' DateTime  : 2003.09.23 12:27
' Author    : Adhimas Setianegara
' Parameters:
'
' Purpose   :
'   If true then the lookup button only displayed when the
'   textbox got the focus
'---------------------------------------------------------------------------------------
'
Public Property Get LookupButtonOnFocus() As Boolean
Attribute LookupButtonOnFocus.VB_Description = "Get/set whether the lookup button should be displayed only when the control got focused."

    LookupButtonOnFocus = mbLookupButtonOnFocus

End Property

Public Property Let LookupButtonOnFocus(ByVal abLookupButtonOnFocus As Boolean)
    mbLookupButtonOnFocus = abLookupButtonOnFocus
    
    If mbLookupButtonOnFocus Then
        SetLocked False
        ShowLookupButton False
    End If
    
    Call UserControl.PropertyChanged("LookupButtonOnFocus")

End Property


'---------------------------------------------------------------------------------------
' Procedure : LookupKey
' DateTime  : 2003.09.23 12:31
' Author    : Adhimas Setianegara
' Parameters:
'   aiLookupKey: vbKey* (vbKeyF3, vbKeyF4 (default), etc.)
' Purpose   :
'   Set keyboard shortcut to click the lookup button
'---------------------------------------------------------------------------------------
'
Public Property Get LookupKey() As Integer
Attribute LookupKey.VB_Description = "Get/set the keyboard shortcut to click the lookup button (vbKeyF3, vbKeyF4, etc.)"

    LookupKey = miLookupKey

End Property

Public Property Let LookupKey(ByVal aiLookupKey As Integer)

    miLookupKey = aiLookupKey

    Call UserControl.PropertyChanged("LookupKey")

End Property

'---------------------------------------------------------------------------------------
' Procedure : LookupKeyShift
' DateTime  : 2003.09.23 12:31
' Author    : Adhimas Setianegara
' Parameters:
'   aiLookupKeyShift: vbAltMask OR vbCtrlMask OR vbShiftMask
' Purpose   :
'   Set keyboard shift-key shortcut to click the lookup button
'---------------------------------------------------------------------------------------
'
Public Property Get LookupKeyShift() As Integer
Attribute LookupKeyShift.VB_Description = "Get/set the keyboard shift-key shortcut to click the lookup button (vbAltMask OR vbCtrlMask OR vbShiftMask.)"

    LookupKeyShift = miLookupKeyShift

End Property

Public Property Let LookupKeyShift(ByVal aiLookupKeyShift As Integer)

    miLookupKeyShift = aiLookupKeyShift

    Call UserControl.PropertyChanged("LookupKeyShift")

End Property

'---------------------------------------------------------------------------------------
' Procedure : LookupButtonInside
' DateTime  : 2003.09.23 13:56
' Author    : Adhimas Setianegara
' Parameters:
' Ret. value:
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get LookupButtonInside() As Boolean
Attribute LookupButtonInside.VB_Description = "Get/set whether the lookup button should be placed on top of the textbox or by extending the control width."

    LookupButtonInside = mbLookupButtonInside

End Property

'---------------------------------------------------------------------------------------
' Procedure : LookupButtonInside
' DateTime  : 2003.09.23 13:56
' Author    : Adhimas Setianegara
' Parameters:
' Ret. value:
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let LookupButtonInside(ByVal abLookupButtonInside As Boolean)
    mbLookupButtonInside = abLookupButtonInside
    ShowLookupButton mbLookupButton
    
    Call UserControl.PropertyChanged("LookupButtonInside")

End Property


'---------------------------------------------------------------------------------------
' Procedure : ShowLookupButton
' DateTime  : 2003.09.23 14:09
' Author    : Adhimas Setianegara
' Parameters:
'
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ShowLookupButton(Optional abStatus As Boolean = True)
    If Not txtForm.Locked Then
        If (abStatus = True) And (mbLookupButton = False) Then
            If Not mbLookupButtonInside Then UserControl.Width = _
                UserControl.Width + cmdForm.Width
            cmdForm.Left = UserControl.Width - cmdForm.Width
        
        ElseIf abStatus = False And mbLookupButton = True Then
            If Not mbLookupButtonInside Then UserControl.Width = _
                UserControl.Width - cmdForm.Width
        End If
        
        cmdForm.Visible = abStatus
    End If
    
    mbLookupButton = abStatus
    ResizeTextBox
    ResizeLabel
End Sub

Private Sub SetFont(aoFont As IFontDisp)
    Set lblAutoSizer.Font = aoFont
    Set txtForm.Font = aoFont
    Set lblForm.Font = aoFont
    
    lblForm.Height = lblAutoSizer.Height
    cmdForm.Height = lblAutoSizer.Height + OnePixelY
    txtForm.Height = lblAutoSizer.Height + OnePixelY
    
    SetStyle
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = lblForm.BackColor
End Property

Public Property Let BackColor(anColor As OLE_COLOR)
    lblForm.BackColor = anColor
    If mbUseContainerColor Then mbUseContainerColor = False
    PropertyChanged "BackColor"
End Property

Public Property Get UseContainerColor() As Boolean

    UseContainerColor = mbUseContainerColor

End Property

Public Property Let UseContainerColor( _
    ByVal bUseContainerColor As Boolean)

    mbUseContainerColor = bUseContainerColor
    If bUseContainerColor Then lblForm.BackColor = _
        UserControl.Parent.BackColor
    
    Call UserControl.PropertyChanged("UseContainerColor")

End Property


Public Property Get EntryBackColor() As OLE_COLOR

    EntryBackColor = mnEntryBackColor

End Property

Public Property Let EntryBackColor(anEntryColor As OLE_COLOR)

    UseContainerColor = False
    mnEntryBackColor = anEntryColor
    ColorizeEntry
    
    Call UserControl.PropertyChanged("EntryBackColor")
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblForm.ForeColor
End Property

Public Property Let ForeColor(anColor As OLE_COLOR)
    If Not mbUseContainerColor Then
        lblForm.ForeColor = anColor
        PropertyChanged "ForeColor"
    End If
End Property

Public Property Get EntryForeColor() As OLE_COLOR
    EntryForeColor = mnEntryForeColor
End Property

Public Property Let EntryForeColor(anColor As OLE_COLOR)
    mnEntryForeColor = anColor
    ColorizeEntry
    PropertyChanged "EntryForeColor"
End Property

Private Sub SetEntryBackColor(anNewColor As OLE_COLOR)
    txtForm.BackColor = anNewColor
    lineLabel.BorderColor = anNewColor
End Sub

Private Sub ColorizeEntry()
    If txtForm.Locked Then
        txtForm.ForeColor = mnLockedForeColor
        SetEntryBackColor mnLockedBackColor
    Else
        txtForm.ForeColor = mnEntryForeColor
        SetEntryBackColor mnEntryBackColor
    End If
End Sub

Public Property Get Locked() As Boolean
    Locked = txtForm.Locked
End Property

Public Property Let Locked(ByVal abLocked As Boolean)

    SetLocked abLocked
    Call UserControl.PropertyChanged("Locked")
End Property

Private Sub SetLocked(abLocked As Boolean)
    If abLocked Then
        ShowLookupButton False
        mbLookupButtonOnFocus = False
    Else
        ShowLookupButton mbLookupButton
    End If
    
    txtForm.Locked = abLocked
    ColorizeEntry
End Sub


Public Property Get LockedBackColor() As OLE_COLOR
    LockedBackColor = mnLockedBackColor
End Property

Public Property Let LockedBackColor(ByVal anLockedBackColor As OLE_COLOR)
    mnLockedBackColor = anLockedBackColor
    ColorizeEntry
    
    Call UserControl.PropertyChanged("LockedBackColor")
End Property

Public Property Get LockedForeColor() As OLE_COLOR
    LockedForeColor = mnLockedForeColor
End Property

Public Property Let LockedForeColor(ByVal anLockedForeColor As OLE_COLOR)
    mnLockedForeColor = anLockedForeColor
    ColorizeEntry
    
    Call UserControl.PropertyChanged("LockedForeColor")
End Property

Public Property Get FocusedBackColor() As OLE_COLOR
    FocusedBackColor = mnFocusedBackColor
End Property

Public Property Let FocusedBackColor(ByVal anFocusedBackColor As OLE_COLOR)
    mnFocusedBackColor = anFocusedBackColor
    Call UserControl.PropertyChanged("FocusedBackColor")
End Property

Public Property Get TextBoxLeft() As Integer
    TextBoxLeft = mnTextBoxLeft
End Property

Public Property Let TextBoxLeft(ByVal anTextBoxLeft As Integer)
    mnTextBoxLeft = anTextBoxLeft
    ResizeTextBox
    ResizeLabel
    Call UserControl.PropertyChanged("TextBoxLeft")
End Property


Public Property Get Style() As LabelStyle

    Style = meStyle

End Property

Public Property Let Style(ByVal eStyle As LabelStyle)

    meStyle = eStyle
    SetStyle
    If Not mbLookupButtonOnFocus Then ShowLookupButton
    
    Call UserControl.PropertyChanged("Style")

End Property


Private Sub SetStyle()
    Select Case meStyle
        Case ltsLabelLeft: ArrangeLabelLeft
        Case ltsLabelTop: ArrangeLabelTop
        Case ltsLabelBottom: ArrangeLabelBottom
    End Select
End Sub

Private Sub ArrangeLabelLeft()
    lblForm.Move 0, 0, lblForm.Width, lblAutoSizer.Height
    txtForm.Move mnTextBoxLeft, 0
    cmdForm.Top = 0

    lineLabel.BorderWidth = 1
    lineLabel.X1 = 0
    lineLabel.X2 = mnTextBoxLeft
    lineLabel.Y1 = lblAutoSizer.Height
    lineLabel.Y2 = lineLabel.Y1

    mnProperHeight = lineLabel.Y2 + OnePixelY
    UserControl_Resize
End Sub

Private Sub ArrangeLabelTop()
    lblForm.Move OFFSET, 0, lblForm.Width, _
        lblAutoSizer.Height + OnePixelY
    txtForm.Move OFFSET, _
        lblAutoSizer.Height + 2 * OnePixelY
    cmdForm.Top = txtForm.Top

    lineLabel.BorderWidth = 4
    lineLabel.X1 = 0
    lineLabel.X2 = 0
    lineLabel.Y1 = 0
    lineLabel.Y2 = txtForm.Top + txtForm.Height

    mnProperHeight = lineLabel.Y2
    UserControl_Resize
End Sub

Private Sub ArrangeLabelBottom()
    ResizeLabel
    
    lblForm.Move OFFSET, txtForm.Height + OnePixelY, _
        lblForm.Width, lblAutoSizer.Height + OnePixelY
    txtForm.Move OFFSET, 0
    cmdForm.Top = txtForm.Top
    
    lineLabel.BorderWidth = 4
    lineLabel.X1 = 0
    lineLabel.X2 = 0
    lineLabel.Y1 = 0
    lineLabel.Y2 = lblForm.Top + lblForm.Height

    mnProperHeight = lineLabel.Y2
    UserControl_Resize
End Sub


Private Function OnePixelY() As Integer
    OnePixelY = Screen.TwipsPerPixelY
End Function

Private Function OnePixelX() As Integer
    OnePixelX = Screen.TwipsPerPixelX
End Function

Private Sub ResizeLabel()
    If (meStyle = ltsLabelLeft) Then
        lblForm.Width = mnTextBoxLeft - OnePixelY
    Else
        If mbLookupButton And mbLookupButtonInside Then
            lblForm.Width = txtForm.Width + cmdForm.Width
        Else
            lblForm.Width = txtForm.Width
        End If
    End If
End Sub
