VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Tutorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tutorial"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Demo"
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7065
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12462
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   15
      Tab             =   14
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1"
      TabPicture(0)   =   "Tutorial.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image14"
      Tab(0).Control(1)=   "Label47"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "2"
      TabPicture(1)   =   "Tutorial.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label10"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "3"
      TabPicture(2)   =   "Tutorial.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(1)=   "Label24"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(5)=   "Label14"
      Tab(2).Control(6)=   "Image1"
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(8)=   "Label21"
      Tab(2).Control(9)=   "Label20"
      Tab(2).Control(10)=   "Label19"
      Tab(2).Control(11)=   "Label18"
      Tab(2).Control(12)=   "Label17"
      Tab(2).Control(13)=   "Label13"
      Tab(2).Control(14)=   "Label12"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "4"
      TabPicture(3)   =   "Tutorial.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label30"
      Tab(3).Control(1)=   "Label29"
      Tab(3).Control(2)=   "Label28"
      Tab(3).Control(3)=   "Image4"
      Tab(3).Control(4)=   "Image3"
      Tab(3).Control(5)=   "Image2"
      Tab(3).Control(6)=   "Label27"
      Tab(3).Control(7)=   "Label26"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "5"
      TabPicture(4)   =   "Tutorial.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image9"
      Tab(4).Control(1)=   "Label46"
      Tab(4).Control(2)=   "Label45"
      Tab(4).Control(3)=   "Label44"
      Tab(4).Control(4)=   "Label43"
      Tab(4).Control(5)=   "Label42"
      Tab(4).Control(6)=   "Label41"
      Tab(4).Control(7)=   "Label40"
      Tab(4).Control(8)=   "Label39"
      Tab(4).Control(9)=   "Label38"
      Tab(4).Control(10)=   "Label37"
      Tab(4).Control(11)=   "Label36"
      Tab(4).Control(12)=   "Label35"
      Tab(4).Control(13)=   "Label34"
      Tab(4).Control(14)=   "Label33"
      Tab(4).Control(15)=   "Label32"
      Tab(4).Control(16)=   "Label31"
      Tab(4).ControlCount=   17
      TabCaption(5)   =   "6"
      TabPicture(5)   =   "Tutorial.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label147"
      Tab(5).Control(1)=   "Label146"
      Tab(5).Control(2)=   "Label145"
      Tab(5).Control(3)=   "Label144"
      Tab(5).Control(4)=   "Label143"
      Tab(5).Control(5)=   "Label142"
      Tab(5).Control(6)=   "Label141"
      Tab(5).Control(7)=   "Label140"
      Tab(5).Control(8)=   "Label139"
      Tab(5).Control(9)=   "Label138"
      Tab(5).Control(10)=   "Label137"
      Tab(5).Control(11)=   "Label136"
      Tab(5).Control(12)=   "Label135"
      Tab(5).Control(13)=   "Label134"
      Tab(5).Control(14)=   "Label133"
      Tab(5).Control(15)=   "Label132"
      Tab(5).Control(16)=   "Image13"
      Tab(5).ControlCount=   17
      TabCaption(6)   =   "7"
      TabPicture(6)   =   "Tutorial.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image15"
      Tab(6).Control(1)=   "Label148"
      Tab(6).Control(2)=   "Label149"
      Tab(6).Control(3)=   "Label150"
      Tab(6).Control(4)=   "Label151"
      Tab(6).Control(5)=   "Label152"
      Tab(6).Control(6)=   "Label153"
      Tab(6).Control(7)=   "Label154"
      Tab(6).Control(8)=   "Label155"
      Tab(6).Control(9)=   "Label156"
      Tab(6).Control(10)=   "Label157"
      Tab(6).Control(11)=   "Label158"
      Tab(6).Control(12)=   "Label159"
      Tab(6).Control(13)=   "Label160"
      Tab(6).Control(14)=   "Label161"
      Tab(6).Control(15)=   "Label162"
      Tab(6).Control(16)=   "Label163"
      Tab(6).ControlCount=   17
      TabCaption(7)   =   "8"
      TabPicture(7)   =   "Tutorial.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label165"
      Tab(7).Control(1)=   "Line6"
      Tab(7).Control(2)=   "Line5"
      Tab(7).Control(3)=   "Line4"
      Tab(7).Control(4)=   "Line3"
      Tab(7).Control(5)=   "Line2"
      Tab(7).Control(6)=   "Line1"
      Tab(7).Control(7)=   "Label164"
      Tab(7).Control(8)=   "Image16"
      Tab(7).ControlCount=   9
      TabCaption(8)   =   "9"
      TabPicture(8)   =   "Tutorial.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Image5"
      Tab(8).Control(1)=   "Label48"
      Tab(8).Control(2)=   "Image6"
      Tab(8).Control(3)=   "Label49"
      Tab(8).Control(4)=   "Label50"
      Tab(8).Control(5)=   "Image7"
      Tab(8).Control(6)=   "Label52"
      Tab(8).Control(7)=   "Label53"
      Tab(8).Control(8)=   "Label54"
      Tab(8).Control(9)=   "Label55"
      Tab(8).Control(10)=   "Label56"
      Tab(8).Control(11)=   "Label57"
      Tab(8).Control(12)=   "Label58"
      Tab(8).ControlCount=   13
      TabCaption(9)   =   "10"
      TabPicture(9)   =   "Tutorial.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label59"
      Tab(9).Control(1)=   "Label60"
      Tab(9).Control(2)=   "Label61"
      Tab(9).Control(3)=   "Label51"
      Tab(9).Control(4)=   "Label62"
      Tab(9).Control(5)=   "Label63"
      Tab(9).Control(6)=   "Label64"
      Tab(9).Control(7)=   "Label65"
      Tab(9).Control(8)=   "Label66"
      Tab(9).Control(9)=   "Label68"
      Tab(9).Control(10)=   "Label67"
      Tab(9).Control(11)=   "Label69"
      Tab(9).Control(12)=   "Label70"
      Tab(9).Control(13)=   "Label71"
      Tab(9).Control(14)=   "Image10"
      Tab(9).ControlCount=   15
      TabCaption(10)  =   "11"
      TabPicture(10)  =   "Tutorial.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Label72"
      Tab(10).Control(1)=   "Label73"
      Tab(10).Control(2)=   "Label74"
      Tab(10).Control(3)=   "Label75"
      Tab(10).Control(4)=   "Label76"
      Tab(10).Control(5)=   "Label77"
      Tab(10).Control(6)=   "Label78"
      Tab(10).Control(7)=   "Label79"
      Tab(10).Control(8)=   "Label80"
      Tab(10).Control(9)=   "Label81"
      Tab(10).Control(10)=   "Label82"
      Tab(10).Control(11)=   "Label83"
      Tab(10).Control(12)=   "Label84"
      Tab(10).Control(13)=   "Label85"
      Tab(10).ControlCount=   14
      TabCaption(11)  =   "12"
      TabPicture(11)  =   "Tutorial.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Label86"
      Tab(11).Control(1)=   "Label94"
      Tab(11).Control(2)=   "Label95"
      Tab(11).Control(3)=   "Label96"
      Tab(11).Control(4)=   "Label97"
      Tab(11).Control(5)=   "Label98"
      Tab(11).Control(6)=   "Label99"
      Tab(11).Control(7)=   "Label100"
      Tab(11).Control(8)=   "Label101"
      Tab(11).Control(9)=   "Label102"
      Tab(11).Control(10)=   "Label103"
      Tab(11).Control(11)=   "Label104"
      Tab(11).Control(12)=   "Label105"
      Tab(11).Control(13)=   "Label106"
      Tab(11).Control(14)=   "Label107"
      Tab(11).Control(15)=   "Label108"
      Tab(11).Control(16)=   "Label109"
      Tab(11).Control(17)=   "Image8"
      Tab(11).Control(18)=   "Label110"
      Tab(11).Control(19)=   "Label111"
      Tab(11).ControlCount=   20
      TabCaption(12)  =   "13"
      TabPicture(12)  =   "Tutorial.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Label93"
      Tab(12).Control(1)=   "Label92"
      Tab(12).Control(2)=   "Label91"
      Tab(12).Control(3)=   "Label90"
      Tab(12).Control(4)=   "Label89"
      Tab(12).Control(5)=   "Label88"
      Tab(12).Control(6)=   "Label87"
      Tab(12).Control(7)=   "Label112"
      Tab(12).Control(8)=   "Label113"
      Tab(12).Control(9)=   "Label114"
      Tab(12).Control(10)=   "Label115"
      Tab(12).Control(11)=   "Label116"
      Tab(12).Control(12)=   "Label117"
      Tab(12).Control(13)=   "Label118"
      Tab(12).Control(14)=   "Label119"
      Tab(12).Control(15)=   "Label120"
      Tab(12).Control(16)=   "Image11"
      Tab(12).ControlCount=   17
      TabCaption(13)  =   "14"
      TabPicture(13)  =   "Tutorial.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Label131"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).Control(1)=   "Label130"
      Tab(13).Control(1).Enabled=   0   'False
      Tab(13).Control(2)=   "Image12"
      Tab(13).Control(2).Enabled=   0   'False
      Tab(13).Control(3)=   "Label129"
      Tab(13).Control(3).Enabled=   0   'False
      Tab(13).Control(4)=   "Label128"
      Tab(13).Control(4).Enabled=   0   'False
      Tab(13).Control(5)=   "Label127"
      Tab(13).Control(5).Enabled=   0   'False
      Tab(13).Control(6)=   "Label126"
      Tab(13).Control(6).Enabled=   0   'False
      Tab(13).Control(7)=   "Label125"
      Tab(13).Control(7).Enabled=   0   'False
      Tab(13).Control(8)=   "Label124"
      Tab(13).Control(8).Enabled=   0   'False
      Tab(13).Control(9)=   "Label121"
      Tab(13).Control(9).Enabled=   0   'False
      Tab(13).Control(10)=   "Label123"
      Tab(13).Control(10).Enabled=   0   'False
      Tab(13).Control(11)=   "Label122"
      Tab(13).Control(11).Enabled=   0   'False
      Tab(13).ControlCount=   12
      TabCaption(14)  =   "15"
      TabPicture(14)  =   "Tutorial.frx":0188
      Tab(14).ControlEnabled=   -1  'True
      Tab(14).Control(0)=   "Label166"
      Tab(14).Control(0).Enabled=   0   'False
      Tab(14).Control(1)=   "Label167"
      Tab(14).Control(1).Enabled=   0   'False
      Tab(14).Control(2)=   "Label168"
      Tab(14).Control(2).Enabled=   0   'False
      Tab(14).Control(3)=   "Label169"
      Tab(14).Control(3).Enabled=   0   'False
      Tab(14).Control(4)=   "Label170"
      Tab(14).Control(4).Enabled=   0   'False
      Tab(14).Control(5)=   "Label171"
      Tab(14).Control(5).Enabled=   0   'False
      Tab(14).Control(6)=   "Label172"
      Tab(14).Control(6).Enabled=   0   'False
      Tab(14).Control(7)=   "Label173"
      Tab(14).Control(7).Enabled=   0   'False
      Tab(14).Control(8)=   "Label174"
      Tab(14).Control(8).Enabled=   0   'False
      Tab(14).ControlCount=   9
      Begin VB.Label Label174 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Vote in Planet Source Code if You Liked This....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3240
         TabIndex        =   177
         Top             =   6480
         Width           =   4800
      End
      Begin VB.Label Label173 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":01A4
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   600
         TabIndex        =   176
         Top             =   5400
         Width           =   9765
      End
      Begin VB.Label Label172 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "That Easy!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   5040
         TabIndex        =   175
         Top             =   4200
         Width           =   1125
      End
      Begin VB.Label Label171 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<----- This will read the Field Sex and give the value to the control Text2"
         Height          =   195
         Left            =   4320
         TabIndex        =   174
         Top             =   2280
         Width           =   5025
      End
      Begin VB.Label Label170 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<----- This will read the Field Name and give the value to the control Text1"
         Height          =   195
         Left            =   4320
         TabIndex        =   173
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label169 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text2.text = NewAdoedc.Recorset!Sex"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   172
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label168 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text1.text = NewAdoedc.Recorset!Name"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   171
         Top             =   2040
         Width           =   2925
      End
      Begin VB.Label Label167 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unbounded Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   600
         TabIndex        =   170
         Top             =   240
         Width           =   2130
      End
      Begin VB.Label Label166 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":027F
         Height          =   495
         Left            =   600
         TabIndex        =   169
         Top             =   600
         Width           =   9615
      End
      Begin VB.Label Label165 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Files And ADOEDC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -67560
         TabIndex        =   168
         Top             =   840
         Width           =   2655
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -70800
         X2              =   -70560
         Y1              =   5040
         Y2              =   5280
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -70800
         X2              =   -70560
         Y1              =   5040
         Y2              =   4800
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -70800
         X2              =   -66600
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -69960
         X2              =   -69720
         Y1              =   3480
         Y2              =   3720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -69960
         X2              =   -69720
         Y1              =   3480
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -69960
         X2              =   -65760
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label164 
         BackStyle       =   0  'Transparent
         Caption         =   "Something Interesting!!  Is that you can use Tables in Text Files and SQL Statements as well."
         Height          =   615
         Left            =   -67560
         TabIndex        =   167
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Image Image16 
         Height          =   5685
         Left            =   -74520
         Picture         =   "Tutorial.frx":031D
         Top             =   720
         Width           =   6525
      End
      Begin VB.Label Label122 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   166
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label123 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":5F28
         Height          =   675
         Left            =   -73680
         TabIndex        =   165
         Top             =   960
         Width           =   8115
      End
      Begin VB.Label Label121 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CancelUpdate.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74280
         TabIndex        =   164
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Label124 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":6038
         Height          =   675
         Left            =   -73560
         TabIndex        =   163
         Top             =   3120
         Width           =   8115
      End
      Begin VB.Label Label125 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":614B
         Height          =   675
         Left            =   -73560
         TabIndex        =   162
         Top             =   3960
         Width           =   8115
      End
      Begin VB.Label Label126 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With new ADO"
         Height          =   195
         Left            =   -71280
         TabIndex        =   161
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label127 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdoedc.Cancel"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -69960
         TabIndex        =   160
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label Label128 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With new ADO"
         Height          =   195
         Left            =   -71160
         TabIndex        =   159
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label Label129 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdoedc.CancelUpdate"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -69840
         TabIndex        =   158
         Top             =   5160
         Width           =   1950
      End
      Begin VB.Image Image12 
         Height          =   1725
         Left            =   -67800
         Picture         =   "Tutorial.frx":6251
         Top             =   4440
         Width           =   3000
      End
      Begin VB.Label Label130 
         BackStyle       =   0  'Transparent
         Caption         =   "Note how the new Ado sends a Message if one Error is presented example is when there are no valid operations to Cancel."
         Height          =   495
         Left            =   -73440
         TabIndex        =   157
         Top             =   5880
         Width           =   4935
      End
      Begin VB.Label Label131 
         BackStyle       =   0  'Transparent
         Caption         =   "This message is hard coded,and you dont need to code it in your button or what ever control you are using!!! "
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   -68760
         TabIndex        =   156
         Top             =   6240
         Width           =   3975
      End
      Begin VB.Label Label93 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<---This will Add the Data from All the Current Bounded Control to the Database"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -70320
         TabIndex        =   155
         Top             =   2160
         Width           =   5565
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdoedc.AddNew "
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72360
         TabIndex        =   154
         Top             =   2160
         Width           =   1590
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With new ADO"
         Height          =   195
         Left            =   -73680
         TabIndex        =   153
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normally will be like this :"
         Height          =   195
         Left            =   -73680
         TabIndex        =   152
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "adoRecordset.AddNew [FieldList [, Values]]"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72360
         TabIndex        =   151
         Top             =   1560
         Width           =   3075
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AddNew.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   150
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label87 
         BackStyle       =   0  'Transparent
         Caption         =   "You use the AddNew method to add a new record to a recordset (if the recordset object can be updated)"
         Height          =   375
         Left            =   -74400
         TabIndex        =   149
         Top             =   840
         Width           =   7695
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This method saves changes to the database.here's how to use this method"
         Height          =   195
         Left            =   -74280
         TabIndex        =   148
         Top             =   3000
         Width           =   5295
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74280
         TabIndex        =   147
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label114 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "adoRecordset.Update  [Fields [, Values]]"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72240
         TabIndex        =   146
         Top             =   3720
         Width           =   2865
      End
      Begin VB.Label Label115 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normally will be like this :"
         Height          =   195
         Left            =   -73560
         TabIndex        =   145
         Top             =   3360
         Width           =   1740
      End
      Begin VB.Label Label116 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With new ADO"
         Height          =   195
         Left            =   -73560
         TabIndex        =   144
         Top             =   4080
         Width           =   1065
      End
      Begin VB.Label Label117 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdoedc.Update"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72240
         TabIndex        =   143
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label118 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<---This will Update the Data from All the Current Bounded Control to the Database"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -70440
         TabIndex        =   142
         Top             =   4320
         Width           =   5805
      End
      Begin VB.Label Label119 
         BackStyle       =   0  'Transparent
         Caption         =   "This message is hard coded,and you dont need to code it in your button or what ever control you are using!!! "
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   -68880
         TabIndex        =   141
         Top             =   6240
         Width           =   3975
      End
      Begin VB.Label Label120 
         BackStyle       =   0  'Transparent
         Caption         =   "Note how the new Ado sends a Message if one Error is presented example is when there are no valid Records to Update."
         Height          =   375
         Left            =   -73560
         TabIndex        =   140
         Top             =   5760
         Width           =   4935
      End
      Begin VB.Image Image11 
         Height          =   1500
         Left            =   -67920
         Picture         =   "Tutorial.frx":76CE
         Top             =   4680
         Width           =   3000
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some Recordset Methods.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   300
         Left            =   -74160
         TabIndex        =   139
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<---adAffectCurrent is selected in  new ADO"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -69960
         TabIndex        =   138
         Top             =   4080
         Width           =   3075
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdoedc.Delete"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72000
         TabIndex        =   137
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With new ADO"
         Height          =   195
         Left            =   -73320
         TabIndex        =   136
         Top             =   3840
         Width           =   1065
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normally will be like this :"
         Height          =   195
         Left            =   -73320
         TabIndex        =   135
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label98 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "adoRecordset.Delete  [AffectRecords]"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -72000
         TabIndex        =   134
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74040
         TabIndex        =   133
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label100 
         BackStyle       =   0  'Transparent
         Caption         =   "The delete method deletes the current record (or group of records)"
         Height          =   375
         Left            =   -74040
         TabIndex        =   132
         Top             =   1200
         Width           =   7695
      End
      Begin VB.Label Label101 
         AutoSize        =   -1  'True
         Caption         =   "adOpenForwardOnly.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71400
         TabIndex        =   131
         Top             =   2760
         Width           =   1845
      End
      Begin VB.Label Label102 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(default) Deletes only the current record"
         Height          =   195
         Left            =   -69480
         TabIndex        =   130
         Top             =   2760
         Width           =   2790
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         Caption         =   "adAffectGroup.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71400
         TabIndex        =   129
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Label Label104 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deletes the records that satisfy the current Filter property setting."
         Height          =   195
         Left            =   -69480
         TabIndex        =   128
         Top             =   3000
         Width           =   4530
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         Caption         =   "adAffectAll.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71400
         TabIndex        =   127
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label Label106 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deletes all records."
         Height          =   195
         Left            =   -69480
         TabIndex        =   126
         Top             =   3240
         Width           =   1350
      End
      Begin VB.Label Label107 
         AutoSize        =   -1  'True
         Caption         =   "adAffectAllChapters.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71400
         TabIndex        =   125
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label108 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deletes all chapter records (hierarchical recordsets)"
         Height          =   195
         Left            =   -69480
         TabIndex        =   124
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label109 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AffectRecords Parameter constants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   -70440
         TabIndex        =   123
         Top             =   2400
         Width           =   3030
      End
      Begin VB.Image Image8 
         Height          =   1500
         Left            =   -67920
         Picture         =   "Tutorial.frx":8673
         Top             =   4680
         Width           =   3000
      End
      Begin VB.Label Label110 
         BackStyle       =   0  'Transparent
         Caption         =   "Note how the new Ado sends a Message if one Error is presented example is when there are no more Records to Delete."
         Height          =   375
         Left            =   -73560
         TabIndex        =   122
         Top             =   5760
         Width           =   4935
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   "This message is hard coded,and you dont need to code it in your button or what ever control you are using!!! "
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   -68880
         TabIndex        =   121
         Top             =   6240
         Width           =   3975
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CursorType.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74040
         TabIndex        =   120
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "The CursorType property hold the type of the cursor the recordset uses.Here are the four possible options:"
         Height          =   315
         Left            =   -74040
         TabIndex        =   119
         Top             =   1200
         Width           =   8745
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "adOpenDynamic.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   118
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label75 
         Caption         =   $"Tutorial.frx":96A5
         Height          =   435
         Left            =   -71640
         TabIndex        =   117
         Top             =   1800
         Width           =   6225
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "adOpenKeyset-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   116
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label Label77 
         Caption         =   $"Tutorial.frx":973C
         Height          =   675
         Left            =   -71640
         TabIndex        =   115
         Top             =   2520
         Width           =   6240
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "adOpenStatic.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   114
         Top             =   3240
         Width           =   1290
      End
      Begin VB.Label Label79 
         Caption         =   $"Tutorial.frx":9826
         Height          =   555
         Left            =   -71640
         TabIndex        =   113
         Top             =   3240
         Width           =   6600
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "adOpenForwardOnly.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   112
         Top             =   3840
         Width           =   1845
      End
      Begin VB.Label Label81 
         Caption         =   $"Tutorial.frx":98EB
         Height          =   675
         Left            =   -71640
         TabIndex        =   111
         Top             =   3840
         Width           =   6690
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DataMember.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -73920
         TabIndex        =   110
         Top             =   4800
         Width           =   1485
      End
      Begin VB.Label Label83 
         BackStyle       =   0  'Transparent
         Caption         =   "This property holds the value that you use when binding controls to Data Environments to specify a command object"
         Height          =   195
         Left            =   -73920
         TabIndex        =   109
         Top             =   5160
         Width           =   8745
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DataSource.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -73920
         TabIndex        =   108
         Top             =   5640
         Width           =   1380
      End
      Begin VB.Label Label85 
         BackStyle       =   0  'Transparent
         Caption         =   "This property holds the object that you use when binding controls to Data Environments ."
         Height          =   195
         Left            =   -73920
         TabIndex        =   107
         Top             =   6000
         Width           =   8745
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This new ADO has a Build in Popup; To Jump to Desired Record  (Right Click)"
         Height          =   195
         Left            =   -74280
         TabIndex        =   106
         Top             =   2640
         Width           =   5520
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOVE.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74280
         TabIndex        =   105
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "This method moves to the given position in a recordset.Normally in ADO you will move like this :"
         Height          =   435
         Left            =   -74280
         TabIndex        =   104
         Top             =   1320
         Width           =   5025
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "adoRecordset.Move NumRecords [, Start]"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -73680
         TabIndex        =   103
         Top             =   2040
         Width           =   2985
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LockType.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74280
         TabIndex        =   102
         Top             =   3240
         Width           =   1185
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":99D8
         Height          =   915
         Left            =   -74280
         TabIndex        =   101
         Top             =   3720
         Width           =   8745
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "adLockReadOnly.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74040
         TabIndex        =   100
         Top             =   4800
         Width           =   1590
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Read-Only,you cannot alter the data."
         Height          =   195
         Left            =   -71880
         TabIndex        =   99
         Top             =   4800
         Width           =   2610
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "adLockPessimistic.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74040
         TabIndex        =   98
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label68 
         Caption         =   $"Tutorial.frx":9B91
         Height          =   675
         Left            =   -71880
         TabIndex        =   97
         Top             =   5040
         Width           =   6240
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "adLockOptimistic.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74040
         TabIndex        =   96
         Top             =   5760
         Width           =   1590
      End
      Begin VB.Label Label69 
         Caption         =   "Optimistic locking,record by record.The provideruses optimistic locking,locking records only when you call the Update method."
         Height          =   435
         Left            =   -71880
         TabIndex        =   95
         Top             =   5760
         Width           =   6240
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "adLockBatchOptimistic.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74040
         TabIndex        =   94
         Top             =   6240
         Width           =   2085
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Optimistic batch Updatesrequired for batch update mode as opposed to immediate update mode."
         Height          =   195
         Left            =   -71880
         TabIndex        =   93
         Top             =   6240
         Width           =   6810
      End
      Begin VB.Image Image5 
         Height          =   1500
         Left            =   -68160
         Picture         =   "Tutorial.frx":9C82
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":AD64
         Height          =   495
         Left            =   -74400
         TabIndex        =   92
         Top             =   2280
         Width           =   5295
      End
      Begin VB.Image Image6 
         Height          =   1500
         Left            =   -68160
         Picture         =   "Tutorial.frx":ADF5
         Top             =   3000
         Width           =   3195
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":BFBE
         Height          =   495
         Left            =   -74400
         TabIndex        =   91
         Top             =   4440
         Width           =   5655
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This new Ado prevents you when your Recordset as no Records in It."
         Height          =   195
         Left            =   -74400
         TabIndex        =   90
         Top             =   6000
         Width           =   4920
      End
      Begin VB.Image Image7 
         Height          =   1500
         Left            =   -68040
         Picture         =   "Tutorial.frx":C04F
         Top             =   5040
         Width           =   3195
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EOF (End of File).-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   89
         Top             =   1320
         Width           =   1920
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":D400
         Height          =   495
         Left            =   -74400
         TabIndex        =   88
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOF (Begin of File).-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   87
         Top             =   2880
         Width           =   2115
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":D4AE
         Height          =   1215
         Left            =   -74400
         TabIndex        =   86
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RecordCount.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   85
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "The Recordcount property is an important one,because it holds the number of records in the recorset."
         Height          =   495
         Left            =   -74400
         TabIndex        =   84
         Top             =   5400
         Width           =   5655
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some Recordset Properties.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   300
         Left            =   -74400
         TabIndex        =   83
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label147 
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Excel 2000 && Newer (*.xls) "
         Height          =   195
         Left            =   -69240
         TabIndex        =   82
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label146 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting an Excel Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -72960
         TabIndex        =   81
         Top             =   600
         Width           =   3195
      End
      Begin VB.Label Label145 
         BackStyle       =   0  'Transparent
         Caption         =   "Actually it sounds more advanced that what it really its.There are yust 3 easy steps to make Excel Databases on the Fly .  "
         Height          =   375
         Left            =   -74160
         TabIndex        =   80
         Top             =   1320
         Width           =   8655
      End
      Begin VB.Label Label144 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: Select the Location of the Database in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   79
         Top             =   1920
         Width           =   7335
      End
      Begin VB.Label Label143 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Select the Command Type of the Recordset in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   78
         Top             =   2280
         Width           =   7815
      End
      Begin VB.Label Label142 
         BackStyle       =   0  'Transparent
         Caption         =   "The Command Type Property holds the type of command used to generate the recordset,and may be  "" From Table""  or  ""From SQL"". "
         Height          =   495
         Left            =   -72840
         TabIndex        =   77
         Top             =   2640
         Width           =   6495
      End
      Begin VB.Label Label141 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":D62B
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   -72840
         TabIndex        =   76
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label Label140 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":D70D
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   -72840
         TabIndex        =   75
         Top             =   3720
         Width           =   6375
      End
      Begin VB.Label Label139 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: In your Form (Code) type this."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   74
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label Label138 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.ConnectDB"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   73
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label Label137 
         BackStyle       =   0  'Transparent
         Caption         =   "<-----Use ConnectDB ,This is to tell the control                     that is going to  Connect an Excel Database"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -68880
         TabIndex        =   72
         Top             =   4800
         Width           =   3840
      End
      Begin VB.Label Label136 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem DataGrid1"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   71
         Top             =   6120
         Width           =   2280
      End
      Begin VB.Label Label135 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text1,""Name"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   70
         Top             =   5400
         Width           =   2580
      End
      Begin VB.Label Label134 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text2,""Address"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   69
         Top             =   5640
         Width           =   2730
      End
      Begin VB.Label Label133 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Label1,""Age"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   68
         Top             =   5880
         Width           =   2520
      End
      Begin VB.Label Label132 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":D7B7
         ForeColor       =   &H00000000&
         Height          =   1155
         Left            =   -68880
         TabIndex        =   67
         Top             =   5400
         Width           =   3840
      End
      Begin VB.Image Image13 
         Height          =   720
         Left            =   -74160
         Picture         =   "Tutorial.frx":D8DA
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image9 
         Height          =   720
         Left            =   -74160
         Picture         =   "Tutorial.frx":E7A4
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":F66E
         ForeColor       =   &H00000000&
         Height          =   1155
         Left            =   -68880
         TabIndex        =   66
         Top             =   5400
         Width           =   3840
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Label1,""Age"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   65
         Top             =   5880
         Width           =   2520
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text2,""Address"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   64
         Top             =   5640
         Width           =   2730
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text1,""Name"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   63
         Top             =   5400
         Width           =   2580
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem DataGrid1"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   62
         Top             =   6120
         Width           =   2280
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "<-----Use ConnectDB ,This is to tell the control          that is going to  Connect an Access Database"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -68880
         TabIndex        =   61
         Top             =   4800
         Width           =   3840
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.ConnectDB"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   60
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: In your Form (Code) type this."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   59
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":F791
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   -72840
         TabIndex        =   58
         Top             =   3720
         Width           =   6375
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":F83B
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   -72840
         TabIndex        =   57
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "The Command Type Property holds the type of command used to generate the recordset,and may be  "" From Table""  or  ""From SQL"". "
         Height          =   495
         Left            =   -72840
         TabIndex        =   56
         Top             =   2640
         Width           =   6495
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Select the Command Type of the Recordset in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   55
         Top             =   2280
         Width           =   7815
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: Select the Location of the Database in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   54
         Top             =   1920
         Width           =   7335
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Actually it sounds more advanced that what it really its.There are yust 3 easy steps to make Access Databases on the Fly .  "
         Height          =   375
         Left            =   -74160
         TabIndex        =   53
         Top             =   1320
         Width           =   8655
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting an Access Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -72960
         TabIndex        =   52
         Top             =   600
         Width           =   3390
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Access 2000 && Newer (*.mdb) "
         Height          =   195
         Left            =   -69240
         TabIndex        =   51
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Steps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74220
         TabIndex        =   50
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.- From the Property Pages you can select Acces Database ,Excel Database or Text Database"
         Height          =   195
         Left            =   -74220
         TabIndex        =   49
         Top             =   6600
         Width           =   6765
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.- Right Click the Ado Control and select Properties (SET AT DESIGN TIME)"
         Height          =   195
         Left            =   -74220
         TabIndex        =   48
         Top             =   6360
         Width           =   5460
      End
      Begin VB.Image Image4 
         Height          =   4665
         Left            =   -70680
         Picture         =   "Tutorial.frx":F91D
         Top             =   1440
         Width           =   5340
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   -74340
         Picture         =   "Tutorial.frx":13242
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   3105
         Left            =   -72900
         Picture         =   "Tutorial.frx":14F08
         Top             =   2160
         Width           =   1770
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":26F26
         Height          =   1095
         Left            =   -74400
         TabIndex        =   47
         Top             =   720
         Width           =   9615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting a Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   46
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.- NextRecord: To navigate to the Next location in the recordset"
         Height          =   195
         Left            =   -70560
         TabIndex        =   45
         Top             =   6240
         Width           =   4590
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4.- LastRecord: To navigate to the Last location in the recordset"
         Height          =   195
         Left            =   -70560
         TabIndex        =   44
         Top             =   6480
         Width           =   4530
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.- PreviousRecord: To navigate to the Previous location in the recordset"
         Height          =   195
         Left            =   -70560
         TabIndex        =   43
         Top             =   6000
         Width           =   5160
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.- FirstRecord: To navigate to the First location in the recordset"
         Height          =   195
         Left            =   -70560
         TabIndex        =   42
         Top             =   5760
         Width           =   4500
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This Control is very similar to the ADO control,with 4 Navigation Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71160
         TabIndex        =   41
         Top             =   5280
         Width           =   6180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Image"
         Height          =   195
         Left            =   -74040
         TabIndex        =   40
         Top             =   5040
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   -74520
         Picture         =   "Tutorial.frx":2705E
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Using New Adoedc Ocx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74520
         TabIndex        =   39
         Top             =   4080
         Width           =   2025
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":28D24
         Height          =   495
         Left            =   -74520
         TabIndex        =   38
         Top             =   3480
         Width           =   9615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Using Data Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74520
         TabIndex        =   37
         Top             =   3120
         Width           =   1710
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":28E2D
         Height          =   495
         Left            =   -74520
         TabIndex        =   36
         Top             =   2520
         Width           =   9615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":28EED
         Height          =   735
         Left            =   -74520
         TabIndex        =   35
         Top             =   1800
         Width           =   9615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<---- This is how the New Ado Looks like"
         Height          =   195
         Left            =   -72600
         TabIndex        =   34
         Top             =   4680
         Width           =   2850
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some Considerations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74520
         TabIndex        =   33
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":2900F
         Height          =   1095
         Left            =   -74520
         TabIndex        =   32
         Top             =   840
         Width           =   9615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":29215
         Height          =   1095
         Left            =   -74400
         TabIndex        =   31
         Top             =   840
         Width           =   9615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Getting Started:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74400
         TabIndex        =   30
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "What Are Databases?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74400
         TabIndex        =   29
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Databases organize data for access and manipulation under programmatic or aplication control."
         Height          =   195
         Left            =   -74400
         TabIndex        =   28
         Top             =   2400
         Width           =   6705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Whats ADO?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74400
         TabIndex        =   27
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":29459
         Height          =   555
         Left            =   -74400
         TabIndex        =   26
         Top             =   3240
         Width           =   8565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Whats ADOX?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74400
         TabIndex        =   25
         Top             =   3840
         Width           =   1230
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":29547
         Height          =   795
         Left            =   -74400
         TabIndex        =   24
         Top             =   4200
         Width           =   8565
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Whats ADOEDC OCX?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74400
         TabIndex        =   23
         Top             =   5040
         Width           =   1920
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":29682
         Height          =   795
         Left            =   -74400
         TabIndex        =   22
         Top             =   5400
         Width           =   8565
      End
      Begin VB.Image Image14 
         Height          =   2250
         Left            =   -72240
         Picture         =   "Tutorial.frx":29783
         Top             =   1440
         Width           =   5400
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":51095
         Height          =   1095
         Left            =   -73440
         TabIndex        =   21
         Top             =   5400
         Width           =   8055
      End
      Begin VB.Image Image15 
         Height          =   720
         Left            =   -74160
         Picture         =   "Tutorial.frx":5124D
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label148 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":52117
         ForeColor       =   &H00000000&
         Height          =   1155
         Left            =   -68880
         TabIndex        =   20
         Top             =   5400
         Width           =   3840
      End
      Begin VB.Label Label149 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Label1,""Age"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   19
         Top             =   5880
         Width           =   2520
      End
      Begin VB.Label Label150 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text2,""Address"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   18
         Top             =   5640
         Width           =   2730
      End
      Begin VB.Label Label151 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem Text1,""Name"""
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   17
         Top             =   5400
         Width           =   2580
      End
      Begin VB.Label Label152 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.AddItem DataGrid1"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   16
         Top             =   6120
         Width           =   2280
      End
      Begin VB.Label Label153 
         BackStyle       =   0  'Transparent
         Caption         =   "<-----Use ConnectDB ,This is to tell the control                     that is going to  Connect an Text File Database"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -68880
         TabIndex        =   15
         Top             =   4800
         Width           =   3840
      End
      Begin VB.Label Label154 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NewAdodc1.ConnectDB"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72240
         TabIndex        =   14
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label Label155 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: In your Form (Code) type this."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   13
         Top             =   4560
         Width           =   3495
      End
      Begin VB.Label Label156 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":5223A
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   -72840
         TabIndex        =   12
         Top             =   3960
         Width           =   6375
      End
      Begin VB.Label Label157 
         BackStyle       =   0  'Transparent
         Caption         =   $"Tutorial.frx":522E4
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   -72840
         TabIndex        =   11
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label Label158 
         BackStyle       =   0  'Transparent
         Caption         =   "The Command Type Property holds the type of command used to generate the recordset,and may be  "" From Table""  or  ""From SQL"". "
         Height          =   495
         Left            =   -72840
         TabIndex        =   10
         Top             =   2640
         Width           =   6495
      End
      Begin VB.Label Label159 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Select the Command Type of the Recordset in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73440
         TabIndex        =   9
         Top             =   2280
         Width           =   7590
      End
      Begin VB.Label Label160 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: Select the Location of the Database in the Options Tab in Property Pages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73440
         TabIndex        =   8
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label Label161 
         BackStyle       =   0  'Transparent
         Caption         =   "Actually it sounds more advanced that what it really its.There are yust 3 easy steps to make Text Files Databases on the Fly .  "
         Height          =   375
         Left            =   -74160
         TabIndex        =   7
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label Label162 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting a Text File Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -72960
         TabIndex        =   6
         Top             =   600
         Width           =   3405
      End
      Begin VB.Label Label163 
         AutoSize        =   -1  'True
         Caption         =   "Text Document (*.txt) "
         Height          =   195
         Left            =   -69240
         TabIndex        =   5
         Top             =   600
         Width           =   1545
      End
      Begin VB.Image Image10 
         Height          =   1905
         Left            =   -70200
         Picture         =   "Tutorial.frx":523FF
         Top             =   720
         Width           =   5400
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW ADOEDC OCX (One New Database Solution for Total Database Begginers )"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   7560
      Width           =   5805
   End
End
Attribute VB_Name = "Tutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
SSTab1.Tab = SSTab1.Tab - 1
End Sub

Private Sub Command3_Click()
Unload Me
Demo1.Show
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
End Sub

