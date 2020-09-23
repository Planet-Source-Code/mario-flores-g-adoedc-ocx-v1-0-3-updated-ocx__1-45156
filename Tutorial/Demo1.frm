VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo1 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo1"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   7800
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      DataBase        =   "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
      TableName       =   "Customers"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bounded Controls "
      Height          =   2655
      Left            =   6360
      TabIndex        =   10
      Top             =   600
      Width           =   5295
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3120
         TabIndex        =   11
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
         TabIndex        =   13
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393216
         Locked          =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bounded Controls"
      Height          =   2655
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   870
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Demo1.frx":0000
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   3720
      Width           =   11055
      _ExtentX        =   19500
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   3360
      Picture         =   "Demo1.frx":0018
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX Data Object Easy Data Control """"Access Demostration"""""
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
      TabIndex        =   16
      Top             =   225
      Width           =   5550
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12105
   End
End
Attribute VB_Name = "Demo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'>                       Demo 1 Access Bounding Controls part 1                          <
'>                                                                                       <
'>          There are 2 Easy steps to use this kind of controls in Access databases      <
'>                                                                                       <
'>              1.- Connect the Database using the """ConnectDB""" Function              <
'>              2.- Bound the desired controls to display the data using """AddItem"""   <
'>                                                                                       <
'>                 Remember to change the DatabaseType to "Access" in Propertys          <
'>                                                                                       <
'><><><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><

Private Sub Command1_Click()
Unload Me: Demo2.Show
End Sub

Private Sub Form_Load()

'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.ConnectDB        '<---- This Will Connect the Ado to the Given Parameters in Property Pages

'2nd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.AddItem DataGrid1
NewAdodc1.AddItem Text1, "CustomerID"
NewAdodc1.AddItem Text2, "CompanyName"
NewAdodc1.AddItem Text3, "ContactName"
NewAdodc1.AddItem Text4, "Address"
NewAdodc1.AddItem DataCombo1, "Country"
NewAdodc1.AddItem DataList1, "Country"
 
End Sub

