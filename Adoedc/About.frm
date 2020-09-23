VERSION 5.00
Begin VB.Form Abouts 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form3"
   Picture         =   "About.frx":0000
   ScaleHeight     =   2565
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sistec_de_juarez@hotmail.com"
      ForeColor       =   &H00F4E6DF&
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   2190
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
Attribute VB_Name = "Abouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Unload Me
End Sub

