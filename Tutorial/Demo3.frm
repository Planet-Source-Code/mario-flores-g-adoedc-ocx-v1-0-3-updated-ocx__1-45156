VERSION 5.00
Object = "*\A..\Adoedc\Ado.vbp"
Begin VB.Form Demo3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo3"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ocx.NewAdodc NewAdodc1 
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   4200
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      DataBase        =   "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
      TableName       =   "Order Details"
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Next"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   495
      Index           =   7
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Update"
      Height          =   495
      Index           =   6
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New"
      Height          =   495
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last"
      Height          =   495
      Index           =   4
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous"
      Height          =   495
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Next"
      Height          =   495
      Index           =   2
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bounded Controls"
      Height          =   2655
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
      Begin VB.TextBox Text4 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         DataSource      =   "NewAdodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UnitPrice"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Caption         =   "First"
      Height          =   495
      Index           =   1
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CancelUpdate"
      Height          =   495
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This example Works for Access,Excel,Text Files Databases"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   19
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here's an example on how to use the new Ado Method's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   2520
      TabIndex        =   18
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Demo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'>                             Demo 3 Access Recordset Methods                           <
'>                                                                                       <
'>          There are 2 Easy steps to use this kind of controls in Access databases      <
'>                                                                                       <
'>              1.- Connect the Database using the """ConnectDB""" Function              <
'>              2.- Bound the desired controls to display the data using """AddItem"""   <
'>                                                                                       <
'>                Remember to change the DatabaseType to "Access" in Propertys           <
'>                                                                                       <
'><><><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><


'********************************************************************************************
'   Notice That:
'                        This Example Works With Excel and Text Files as well..
'********************************************************************************************

Private Sub Form_Load()

'1st Step<><><><><><><><><><><><><><><><><><><><>
NewAdodc1.ConnectDB
                    '<---- This Will Connect the Ado to the Given Parameters in Property Pages

'2nd Step<><><><><><><><><><><><><><><><><><><><>

NewAdodc1.AddItem Text1, "ProductID"
NewAdodc1.AddItem Text2, "UnitPrice"
NewAdodc1.AddItem Text3, "Quantity"
NewAdodc1.AddItem Text4, "Discount"
End Sub

Private Sub Command_Click(Index As Integer)

If Index = 0 Then NewAdodc1.CancelUpdate
If Index = 1 Then NewAdodc1.FirstRecord
If Index = 2 Then NewAdodc1.NextRecord
If Index = 3 Then NewAdodc1.PreviousRecord
If Index = 4 Then NewAdodc1.LastRecord
If Index = 5 Then NewAdodc1.AddNew
If Index = 6 Then NewAdodc1.Update        '<----- Recordset Methods can be call like this
If Index = 7 Then NewAdodc1.Delete

End Sub

Private Sub Command9_Click()
Unload Me: Demo4.Show
End Sub

