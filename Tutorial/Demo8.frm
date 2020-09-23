VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo8"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   5880
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      DataBaseType    =   2
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   0
      Left            =   10680
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   1
      Left            =   10680
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   2
      Left            =   10680
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   3
      Left            =   10680
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   4
      Left            =   10680
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   5
      Left            =   10680
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   6
      Left            =   10680
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Try Me!"
      Height          =   375
      Index           =   7
      Left            =   10680
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   6600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "<---- Now you can use Tables in Text Files                 (like in Access and Excel)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   5640
      TabIndex        =   27
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The SELECT statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT * FROM [Names.txt]"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   960
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecting Specific Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT [FirstName],[Age] FROM [Names.txt]"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   1680
      Width           =   3225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The WHERE Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   2100
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT * FROM [Spelling.txt] WHERE [Correct]= ""False"""
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The BETWEEN Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   2385
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT * FROM [Names.txt] WHERE [ID] BETWEEN ""3*"" And ""6*"""
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   4905
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The IN Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT * FROM [Spelling.txt] WHERE Spanish IN (""Azul"",""Rojo"")"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   4725
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The DISTINCT Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   2310
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT DISTINCT  Age FROM [Names.txt]"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Order BY Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   2205
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT * FROM [Names.txt] Order BY [LastName]"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The AS Clause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   1560
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT [Age] AS Edades FROM [Names.txt]"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   6000
      Width           =   3225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "for Text files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Labela 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOME BASIC SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BE6B47&
      Height          =   420
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11880
   End
End
Attribute VB_Name = "Demo8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'>                               Demo 8 Text Files Controls & SQL                        <
'>                                                                                       <
'>                 There are 2 Easy steps to use this in Text Files databases            <
'>                                                                                       <
'>         1.- Connect the Ado Recordsource using the """RecordSource""" Function        <
'>         2.- Bound the desired controls to display the data using """AddItem"""        <
'>                                                                                       <
'>                Remember to change the DatabaseType to "TextFile" in Propertys         <
'>                                                                                       <
'><><><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><

Private Sub Command_Click(Index As Integer)

'<><><><><><><><><><><><><><><><><><><><><><><><>
NewAdodc1.Database = App.Path & "\Databases\Names.txt"
                                  '<----Note that there are no Parameters in the Property Pages
                                       'Because Parameters were manually called
                                       'Note that Method must have Header(Principal) Table..in this case: ""Names.txt""
                                       'This is used to get the Location where the tables are Located

'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.RecordSource Label(Index).Caption

'2nd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.AddItem DataGrid1

End Sub


