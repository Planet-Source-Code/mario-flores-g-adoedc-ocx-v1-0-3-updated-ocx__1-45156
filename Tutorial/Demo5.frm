VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo5 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo5"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   7320
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      DataBaseType    =   1
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Next"
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bounded Controls "
      Height          =   2415
      Left            =   6600
      TabIndex        =   8
      Top             =   720
      Width           =   4575
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataList DataList1 
         Height          =   1425
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393216
         Locked          =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirstName"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LastName"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bounded Controls"
      Height          =   2415
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      Begin VB.TextBox Text4 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Demo5.frx":0000
      Height          =   3855
      Left            =   480
      TabIndex        =   7
      Top             =   3840
      Width           =   7335
      _ExtentX        =   12938
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
      Caption         =   $"Demo5.frx":0018
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Index           =   0
      Left            =   8160
      TabIndex        =   17
      Top             =   4800
      Width           =   3225
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   2760
      Picture         =   "Demo5.frx":00C0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX Data Object Easy Data Control """"Excel Demostration"""""
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
      Left            =   3600
      TabIndex        =   14
      Top             =   240
      Width           =   5400
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12105
   End
End
Attribute VB_Name = "Demo5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'>                                Demo 5 Excel Bounding Controls                            <
'>                                                                                          <
'>          There are 3 Easy steps to use this kind of controls in Excel databases          <
'>             1.- Select the Database Location using the """DataBase""" Function           <                                                                    <
'>             2.- Connect the Database using the """RecordSource""" Function               <
'>             3.- Bound the desired controls to display the data using """AddItem"""       <
'>                                                                                          <
'>                Remember to change the DatabaseType to "Excel" in Propertys               <
'>                                                                                          <
'><>< ><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><><


Private Sub Form_Load()


'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.Database = App.Path & "\Databases\Book1.xls"
                                  '<----Notice that there are no Parameters in the Property Pages
                                       'Because Parameters were manually called

'2nd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.RecordSource "SELECT * FROM [Name$]"   '<---Notice that Table haves a delimiter "$" next to it

'3rd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.AddItem DataGrid1
NewAdodc1.AddItem Text1, "FirstName"
NewAdodc1.AddItem Text2, "LastName"
NewAdodc1.AddItem Text3, "Age"
NewAdodc1.AddItem Text4, "Sex"
NewAdodc1.AddItem DataCombo1, "FirstName"      '    <--- Note that Order is not Important!!!
NewAdodc1.AddItem DataList1, "LastName"
End Sub

Private Sub Command9_Click()
Unload Me: Demo6.Show
End Sub
