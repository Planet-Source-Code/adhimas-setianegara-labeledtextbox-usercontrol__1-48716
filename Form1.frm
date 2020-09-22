VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LabeledTextBox txtName 
      Height          =   210
      Left            =   525
      TabIndex        =   0
      Top             =   405
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   370
      LookupButton    =   0   'False
      Caption         =   "Name:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   0   'False
      LookupKey       =   115
      LookupKeyShift  =   0
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   12632256
      ForeColor       =   -2147483630
      EntryBackColor  =   -2147483643
      EntryForeColor  =   16711680
      LockedBackColor =   -2147483643
      LockedForeColor =   -2147483640
      FocusedBackColor=   -2147483643
      Locked          =   0   'False
      TextBoxLeft     =   1470
      Style           =   0
   End
   Begin Project1.LabeledTextBox txtCompany 
      Height          =   435
      Index           =   0
      Left            =   375
      TabIndex        =   3
      Top             =   1620
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      LookupButton    =   0   'False
      Caption         =   "Borland:"
      MaxLength       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   -1  'True
      LookupKey       =   115
      LookupKeyShift  =   0
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   32896
      ForeColor       =   16777215
      EntryBackColor  =   -2147483643
      EntryForeColor  =   -2147483640
      LockedBackColor =   0
      LockedForeColor =   0
      FocusedBackColor=   16777215
      Locked          =   0   'False
      TextBoxLeft     =   60
      Style           =   1
   End
   Begin Project1.LabeledTextBox txtCompany 
      Height          =   435
      Index           =   1
      Left            =   1725
      TabIndex        =   4
      Top             =   1620
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      LookupButton    =   0   'False
      Caption         =   "Sun:"
      MaxLength       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   -1  'True
      LookupKey       =   0
      LookupKeyShift  =   0
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   32896
      ForeColor       =   16777215
      EntryBackColor  =   -2147483643
      EntryForeColor  =   -2147483640
      LockedBackColor =   -2147483643
      LockedForeColor =   -2147483640
      FocusedBackColor=   16777215
      Locked          =   0   'False
      TextBoxLeft     =   60
      Style           =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2970
      Width           =   1065
   End
   Begin Project1.LabeledTextBox txtCompany 
      Height          =   435
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   1620
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      LookupButton    =   0   'False
      Caption         =   "Microsoft:"
      MaxLength       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   -1  'True
      LookupKey       =   0
      LookupKeyShift  =   0
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   32896
      ForeColor       =   16777215
      EntryBackColor  =   -2147483643
      EntryForeColor  =   -2147483640
      LockedBackColor =   -2147483643
      LockedForeColor =   -2147483640
      FocusedBackColor=   16777215
      Locked          =   0   'False
      TextBoxLeft     =   60
      Style           =   1
   End
   Begin Project1.LabeledTextBox txtCountry 
      Height          =   225
      Left            =   525
      TabIndex        =   2
      Top             =   945
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   370
      LookupButton    =   0   'False
      Caption         =   "Country:"
      Text            =   "USA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   0   'False
      LookupKey       =   32
      LookupKeyShift  =   2
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   12632256
      ForeColor       =   -2147483630
      EntryBackColor  =   -2147483643
      EntryForeColor  =   16711680
      LockedBackColor =   -2147483643
      LockedForeColor =   -2147483631
      FocusedBackColor=   16777215
      Locked          =   -1  'True
      TextBoxLeft     =   1470
      Style           =   0
   End
   Begin Project1.LabeledTextBox txtCity 
      Height          =   225
      Left            =   525
      TabIndex        =   1
      ToolTipText     =   "Press F4 to show lookup contents"
      Top             =   675
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   370
      LookupButton    =   0   'False
      Caption         =   "City:"
      MaxLength       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LookupButtonOnFocus=   -1  'True
      LookupKey       =   115
      LookupKeyShift  =   0
      LookupButtonInside=   0   'False
      UseContainerColor=   0   'False
      BackColor       =   12632256
      ForeColor       =   -2147483630
      EntryBackColor  =   -2147483643
      EntryForeColor  =   16711680
      LockedBackColor =   16761024
      LockedForeColor =   16744576
      FocusedBackColor=   16777215
      Locked          =   0   'False
      TextBoxLeft     =   1470
      Style           =   0
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "You should modify the Z-ordering of the controls to display their lookup button properly."
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   525
      TabIndex        =   6
      Top             =   2295
      Width           =   3540
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpNote 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   375
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3840
   End
   Begin VB.Shape frame 
      BackColor       =   &H00C0E0FF&
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   3990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    If vbOK = MsgBox("Are you sure you want to exit?", vbQuestion + vbOKCancel) Then Unload Me
End Sub


Private Sub txtCity_ButtonClick()
    Dim sRes As String
    sRes = InputBox("Please enter a city name", "Question")
    
    If Len(sRes) > 0 Then txtCity.Text = sRes
End Sub


Private Sub txtCompany_GotFocus(Index As Integer)
    txtCompany(Index).Style = ltsLabelBottom
End Sub

Private Sub txtCompany_LostFocus(Index As Integer)
    txtCompany(Index).Style = ltsLabelTop
End Sub

Private Sub txtCountry_ButtonClick()
    Dim sRes As String
    sRes = InputBox("Please enter a country name", "Question")
    
    If Len(sRes) > 0 Then txtCountry.Text = sRes
End Sub


Private Sub txtName_ButtonClick()
    MsgBox "Aneh"
End Sub

