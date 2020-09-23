VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo2 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo2"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   6600
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Demo2.frx":0000
      Top             =   4320
      Width           =   3255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16049887
      BackColorFixed  =   16250871
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12105
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Now! is Easy to Use Access Databases with yust some easy codet!"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5880
      TabIndex        =   13
      Top             =   5280
      Width           =   4785
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   9600
      TabIndex        =   10
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   7320
      TabIndex        =   9
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   8
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Label"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   9600
      TabIndex        =   5
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Label"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   7320
      TabIndex        =   4
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Label"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4920
      TabIndex        =   3
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Label"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Label"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   1020
   End
End
Attribute VB_Name = "Demo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'                        Demo 2 Access Bounding Controls part 2                             <
'>                                                                                          <
'>          There are 3 Easy steps to use this kind of controls in Access databases         <
'>             1.- Select the Database Location using the """DataBase""" Function           <                                                                    <
'>             2.- Connect the Database using the """RecordSource""" Function               <
'>             3.- Bound the desired controls to display the data using """AddItem"""       <
'>                                                                                          <
'>                 Remember to change the DatabaseType to "Access" in Propertys             <
'>                                                                                          <
'><>< ><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><><



'********************************************************************************************
'       This Example Demostrates how to Connect Controls in one Diferent way by
'       Bounding the Controls to the Database, selecting the parameters in code.
'********************************************************************************************

Private Sub Form_Load()

'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.Database = "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
                                  '<----Note that there are no Parameters in the Property Pages
                                  '     And where are not using ConnectDB ...
'2nd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.RecordSource "Select * FROM Employees"

'3rd Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.AddItem MSHFlexGrid1
NewAdodc1.AddItem Label(0), "EmployeeID"
NewAdodc1.AddItem Label(1), "FirstName"
NewAdodc1.AddItem Label(2), "LastName"
NewAdodc1.AddItem Label(3), "Title"
NewAdodc1.AddItem Label(4), "City"                 '    <--- Note that Order is not Important!!!
NewAdodc1.AddItem Text1, "Notes"

End Sub

Private Sub Command1_Click()
Unload Me: Demo3.Show
End Sub

