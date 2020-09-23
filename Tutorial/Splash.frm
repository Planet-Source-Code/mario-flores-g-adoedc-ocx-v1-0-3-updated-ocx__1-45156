VERSION 5.00
Begin VB.Form splash 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form3"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   2565
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2565
      Left            =   0
      Top             =   0
      Width           =   5400
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ADOEDC OCX v1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   0
      Top             =   2280
      Width           =   1410
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Time

Private Sub Timer1_Timer()
Time = Time + 1
If Time < 3 Then Exit Sub
Tutorial.Show
Unload Me
End Sub
