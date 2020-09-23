VERSION 5.00
Begin VB.Form Pop 
   BackColor       =   &H00BE6B47&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F4FDFF&
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   800
      Width           =   735
   End
   Begin VB.CommandButton BotonOk 
      BackColor       =   &H00F4FDFF&
      Caption         =   "Ok"
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "  Goto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   3
      Top             =   15
      Width           =   3990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Record?"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   405
      Width           =   1395
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4FDFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4020
   End
End
Attribute VB_Name = "Pop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Tramp As Boolean

Private Sub BotonOk_Click()
On Error Resume Next
Indice = Abs(CInt(Text1)) + 3
Unload Me
End Sub

'Private Sub BotonOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Tramp = True
'End Sub
'Private Sub BotonOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Tramp = False
'Text1.SetFocus
'End Sub
'Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Tramp = True
'End Sub

'Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Tramp = False

'End Sub

Private Sub Command2_Click()
Indice = -1
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Create_Tooltip TTBalloon, Text1, "Type here the Number of the Record", TTIconInfo, "To Move", vbBlack, vbWhite
End Sub

'Private Sub Text1_LostFocus()
'If Tramp = False Then Unload Me
'End Sub
