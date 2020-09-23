VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo7 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo7"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   7440
      TabIndex        =   18
      Top             =   7560
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      DataBaseType    =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bounded Controls"
      Height          =   2655
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age "
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1680
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bounded Controls "
      Height          =   2655
      Left            =   6360
      TabIndex        =   2
      Top             =   600
      Width           =   5295
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataList DataList1 
         Height          =   1425
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393216
         Locked          =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LastName"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirstName"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   7800
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Demo7.frx":0000
      Height          =   3855
      Left            =   360
      TabIndex        =   13
      Top             =   4080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   104.882
         EndProperty
      EndProperty
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notice that all 3 Databases Options and Methods (Access,Excel,TextFiles) are used  and coded almost the same way .."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   675
      Index           =   1
      Left            =   7560
      TabIndex        =   17
      Top             =   6000
      Width           =   4065
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now! is Easy to Use TextFiles Databases  Yust use the same Ado and select one Text File Database the Ado will do the rest!"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Index           =   0
      Left            =   7920
      TabIndex        =   16
      Top             =   4200
      Width           =   3225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX Data Object Easy Data Control """"Text Files Demostration"""""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3960
      TabIndex        =   14
      Top             =   225
      Width           =   5760
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   3360
      Picture         =   "Demo7.frx":0018
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12105
   End
End
Attribute VB_Name = "Demo7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'>                          Demo 7 Text Files Bounding Controls                          <
'>                                                                                       <
'>       There are 2 Easy steps to use this kind of controls in Text Files databases     <
'>                                                                                       <
'>              1.- Connect the Database using the """ConnectDB""" Function              <
'>              2.- Bound the desired controls to display the data using """AddItem"""   <
'>                                                                                       <
'>                Remember to change the DatabaseType to "TextFile" in Propertys         <
'>                                                                                       <
'><><><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><


Private Sub Form_Load()

'<><><><><><><><><><><><><><><><><><><><><><><><>
NewAdodc1.Database = App.Path & "\Databases\Names.txt"
                                  '<----Notice that there are no Parameters in the Property Pages
                                       'Because Parameters were manually called

'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.RecordSource "Select * From [Names.txt]" ' <--Notice that the Name of the table haves ".txt" Next to it..
                                                   '    indicating that it is a Text File Table..it also can go like
                                                   '    this [Names#txt] where "." Or "#" is the delimiter

'2nd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.AddItem DataGrid1
NewAdodc1.AddItem DataList1, "FirstName"
NewAdodc1.AddItem DataCombo1, "LastName"
NewAdodc1.AddItem Text1, "FirstName"
NewAdodc1.AddItem Text2, "LastName"
NewAdodc1.AddItem Text3, "Age"

End Sub

Private Sub Command1_Click()
Unload Me: Demo8.Show
End Sub

