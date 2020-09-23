VERSION 5.00
Begin VB.UserControl NewAdodc 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   DataSourceBehavior=   1  'vbDataSource
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "UserControl1.ctx":0010
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   7
      Left            =   1320
      Picture         =   "UserControl1.ctx":0322
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   11
      Top             =   2400
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   6
      Left            =   960
      Picture         =   "UserControl1.ctx":0A24
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   10
      Top             =   2400
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   5
      Left            =   600
      Picture         =   "UserControl1.ctx":1126
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   9
      Top             =   2400
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   4
      Left            =   240
      Picture         =   "UserControl1.ctx":1828
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   8
      Top             =   2400
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   3
      Left            =   1320
      Picture         =   "UserControl1.ctx":1F2A
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   7
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   2
      Left            =   960
      Picture         =   "UserControl1.ctx":262C
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   6
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   1
      Left            =   600
      Picture         =   "UserControl1.ctx":2D2E
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   5
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox ButtonImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   240
      Picture         =   "UserControl1.ctx":3430
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   4
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox ButtonAdo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   3
      Left            =   1080
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   3
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox ButtonAdo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox ButtonAdo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   1
      Left            =   360
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   1
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox ButtonAdo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   2
      Left            =   720
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FF80&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "NewAdodc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'>                                    ADOEDC OCX BUILD 1.0                               <
'>                                                                                       <
'>                                      By MArio Flores G                                <
'>                                                                                       <
'>                           ***This is as a Database Tool Control***                    <
'>                                                                                       <
'>        READ THIS:                                                                     <
'>                                                                                       <
'>                                                                                       <
'>             Commercial use of this control and either any part of this code           <
'>             is FORBIDDEN without explicitly permission from me.                       <
'>             You can use this code for your personal projects, or for freeware         <
'>             but you must inlcude the original unmodified about form.                  <
'>                                                                                       <
'>                                sistec_de_juarez@hotmail.com                           <
'>                                                                                       <
'>                                 Cd Juarez Chihuahua Mexico                            <
'>                                                                                       <
'><><><><><><><><><><><><><><><><>'ADOEDC by MArio Flores G<><><><><><><><><><><><><><><><


Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Running As Boolean

Public Connection As Connection
Public RecordSet As RecordSet

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private i

Public Enum NewType
    DBtables = 0
    SQL = 1
End Enum

Public Enum NewDBType
    Access = 0
    Excel = 1
    TextFile = 2
End Enum


Private CDataBase As String
Private CTableSL As String
Private CTextSQL As String
Private CType As NewType
Private CDBType As NewDBType
Private Const CCDataBase = vbNullString
Private Const CCTableSL = vbNullString
Private Const CCTextSQL = vbNullString
Private Const CCType = NewType.DBtables
Private Const CCDBType = NewDBType.Access
Private Pt As POINTAPI



Private Sub ButtonAdo_Click(Index As Integer)
If Index = 0 Then FirstRecord
If Index = 1 Then PreviousRecord
If Index = 2 Then NextRecord
If Index = 3 Then LastRecord
If Running = False Then Create_Tooltip TTBalloon, ButtonAdo(Index), "Check RecordSource", TTIconWarning, "ADO not Connected!", vbBlack, vbWhite
End Sub


Private Sub ButtonAdo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = vbRightButton Then
   GetCursorPos Pt
    Pop.Move Pt.x * Screen.TwipsPerPixelX, (Pt.y - 80) * Screen.TwipsPerPixelY
    Pop.Show 1
    If Indice <> -1 Then MoveRecords Indice
    Exit Sub
End If

BitBlt ButtonAdo(Index).hdc, 0, 0, 24, 24, ButtonImage(Index + 4).hdc, 0, 0, vbSrcCopy

End Sub

Private Sub ButtonAdo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
BitBlt ButtonAdo(Index).hdc, 0, 0, 24, 24, ButtonImage(Index).hdc, 0, 0, vbSrcCopy
End Sub

'--------------------------------------------------------------------
'--------------------------------------------------------------------
'                         PROPERTYS DECLARATIONS

   Public Property Get DataBase() As String
    DataBase = CDataBase
   End Property
                                                                    
   Public Property Let DataBase(ByVal NewValue As String)
    CDataBase = NewValue
    PropertyChanged "DataBase"
   End Property

   Public Property Get TableName() As String
Attribute TableName.VB_MemberFlags = "40"
    TableName = CTableSL
   End Property

   Public Property Let TableName(ByVal NewValue As String)
    CTableSL = NewValue
    PropertyChanged "TableName"
   End Property
    
   Public Property Get TextSQL() As String
Attribute TextSQL.VB_MemberFlags = "40"
    TextSQL = CTextSQL
   End Property

   Public Property Let TextSQL(ByVal NewValue As String)
    CTextSQL = NewValue
    PropertyChanged "TextSQL"
   End Property
   
   Public Property Get CommandType() As NewType
    CommandType = CType
   End Property

   Public Property Let CommandType(ByVal NewValue As NewType)
    CType = NewValue
    PropertyChanged "CommandType"
   End Property
   
   Public Property Get DataBaseType() As NewDBType
    DataBaseType = CDBType
   End Property

   Public Property Let DataBaseType(ByVal NewValue As NewDBType)
    CDBType = NewValue
    PropertyChanged "DataBaseType"
   End Property
'--------------------------------------------------------------------
'--------------------------------------------------------------------

Private Sub UserControl_Initialize()
Dim px As Long, py As Long
px& = 24
py& = 24

For i = ButtonAdo.LBound To ButtonAdo.UBound
    BitBlt ButtonAdo(i).hdc, 0, 0, 24, 24, ButtonImage(i).hdc, 0, 0, vbSrcCopy
    ButtonAdo(i).AutoRedraw = False
Next i

CDataBase = CCDataBase
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Destroy_Tooltip
End Sub

'--------------------------------------------------------------------
'--------------------------------------------------------------------
'                        PROPERTYBAG READ & WRITE

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
DataBase = PropBag.ReadProperty("DataBase", CCDataBase)
TableName = PropBag.ReadProperty("TableName", CCTableSL)
TextSQL = PropBag.ReadProperty("TextSQL", CCTextSQL)
CommandType = PropBag.ReadProperty("CommandType", CCType)
DataBaseType = PropBag.ReadProperty("DataBaseType", CCDBType)
End Sub


Private Sub UserControl_Resize()
UserControl.Width = 1440
UserControl.Height = 370
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("DataBase", CDataBase, CCDataBase)
Call PropBag.WriteProperty("TableName", CTableSL, CCTableSL)
Call PropBag.WriteProperty("TextSQL", CTextSQL, CCTextSQL)
Call PropBag.WriteProperty("CommandType", CType, CCType)
Call PropBag.WriteProperty("DataBaseType", CDBType, CCDBType)
End Sub


'--------------------------------------------------------------------
'--------------------------------------------------------------------
'PRIVATE SUB CONNECTACCESSDB(To CONNECT ACCESS DATABASE)

Public Sub ConnectDB(Optional Connect As String)
Attribute ConnectDB.VB_MemberFlags = "40"
On Error GoTo XError



Set Connection = New Connection
Set RecordSet = New RecordSet
Destroy_Tooltip

If Len(DataBase) = 0 Then GoTo NoDataBaseError

If DataBaseType = Access Then Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & DataBase & ";Persist Security Info=False"
If DataBaseType = Excel Then Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & DataBase & ";Extended Properties=Excel 8.0;"
If DataBaseType = TextFile Then Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & TextFormatPath(DataBase, 2) & ";Extended Properties=Text;"

Connection.CursorLocation = adUseClient

If Len(Connect) <> 0 Then GoTo R


If CommandType = 0 Then
    
    If Len(TableName) = 0 Then GoTo CommandTypeError
    
    If DataBaseType = TextFile Then
        RecordSet.Open "Select * From [" & TextFormatPath(DataBase, 1) & "]", Connection, adOpenDynamic, adLockOptimistic
    End If
    If DataBaseType = Access Or DataBaseType = Excel Then
        RecordSet.Open "Select * From [" & TableName & "]", Connection, adOpenDynamic, adLockOptimistic
    End If

End If

If CommandType = 1 Then
    If Len(TextSQL) = 0 Then GoTo CommandTypeError
     RecordSet.Open TextSQL, Connection, adOpenDynamic, adLockOptimistic
End If

Running = True
Exit Sub

R:
RecordSet.Open Connect, Connection, adOpenDynamic, adLockOptimistic
Running = True
Exit Sub

'Common Errors in Database^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
NoDataBaseError:
MsgBox "Error Locating Database", vbExclamation, "NewAdodc OCX"
Exit Sub
CommandTypeError:
MsgBox "Error in CommandType", vbExclamation, "NewAdodc OCX"
Exit Sub
XError:
MsgBox Err.Description, vbCritical, "NewAdodc OCX"
End Sub


'--------------------------------------------------------------------
'----------------------   Navigation Buttons  -----------------------

Public Sub FirstRecord()
MoveRecords 0
End Sub
Public Sub PreviousRecord()
MoveRecords 1
End Sub
Public Sub NextRecord()
MoveRecords 2
End Sub
Public Sub LastRecord()
MoveRecords 3
End Sub

'------------------------    ADDITEM  -------------------------------
'--------------------------------------------------------------------
'           (Datasource and Datafield Set of the Controls)

Public Sub AddItem(Names As Object, Optional Field As String)
    On Error Resume Next
    If Running = False Then Exit Sub
    Set Names.DataSource = RecordSet
    Set Names.RowSource = RecordSet
    If Len(Field) <> 0 Then
         Names.BoundColumn = Field
         Names.ListField = Field
         Names.DataField = Field
         Names.Picture = LoadPicture(Field)
    End If
End Sub

'-----------------------    RECORDSOURCE    -------------------------
'--------------------------------------------------------------------
'                    (Recordsource Set of ADO)

Public Sub RecordSource(Command As String)
       ConnectDB Command
End Sub

'-----------------------    MOVE RECORDS    -------------------------
'--------------------------------------------------------------------

Private Sub MoveRecords(Direction As Integer)
On Error GoTo XError

If Running = False Then Exit Sub

With RecordSet
    
    If .RecordCount = 0 Then
        Create_Tooltip TTBalloon, ButtonAdo(Direction), "There are no Records", TTIconInfo, "No Record Found", vbBlack, vbWhite
        Exit Sub
    End If
    
    If Direction = 0 Then .MoveFirst
    If Direction = 1 Then .MovePrevious
    If Direction = 2 Then .MoveNext
    If Direction = 3 Then .MoveLast
    If Direction > 3 Then
    .MoveFirst
     On Local Error Resume Next
    .Move Direction - 4
    End If
    
    If .BOF = True Then
        .MoveNext
        Create_Tooltip TTBalloon, ButtonAdo(Direction), "There are no more Records", TTIconInfo, "BOF", vbBlack, vbWhite
        Exit Sub
    End If
    
    If .EOF = True Then
        .MovePrevious
        Create_Tooltip TTBalloon, ButtonAdo(Direction), "There are no more Records", TTIconInfo, "EOF", vbBlack, vbWhite
        Exit Sub
    End If
    
    Create_Tooltip TTStandard, ButtonAdo(Direction), Abs(.AbsolutePosition), TTIconInfo, , vbBlack, vbWhite

End With


Exit Sub

XError:
MsgBox Err.Description, vbCritical, "NewAdodc OCX"

End Sub

'-------------------------  RECORSET METHODS ------------------------
'--------------------------------------------------------------------
  
Public Sub AddNew()
    On Error Resume Next
    If Running = False Then Exit Sub
    Destroy_Tooltip
    RecordSet.AddNew
End Sub

Public Sub Delete()
    On Error Resume Next
     If Running = False Then Exit Sub
     If RecordSet.BOF = True Or RecordSet.EOF = True Then
        Create_Tooltip TTBalloon, Parent.ActiveControl, "There are no more Records", TTIconError, "Error", vbBlack, vbWhite
        Exit Sub
     End If
     Destroy_Tooltip
     RecordSet.Delete adAffectCurrent
     RecordSet.Requery
End Sub

Public Sub Update()
    On Error GoTo MissingError
    If Running = False Then Exit Sub
    Destroy_Tooltip
    RecordSet.Update
    Exit Sub
MissingError:
    Create_Tooltip TTBalloon, Parent.ActiveControl, "Missing Field", TTIconError, "Can't Update", vbBlack, vbWhite
End Sub

Public Sub Cancel()
    On Error GoTo MissingError
    If Running = False Then Exit Sub
    Destroy_Tooltip
    RecordSet.Cancel
    Exit Sub
MissingError:
    Create_Tooltip TTBalloon, Parent.ActiveControl, "Operation Failed", TTIconError, "Can't Cancel", vbBlack, vbWhite
End Sub

Public Sub CancelUpdate()
    On Error GoTo MissingError
    If Running = False Then Exit Sub
    Destroy_Tooltip
    RecordSet.CancelUpdate
    Exit Sub
MissingError:
    Create_Tooltip TTBalloon, Parent.ActiveControl, "Operation Failed", TTIconError, "Can't Cancel Update", vbBlack, vbWhite
End Sub

Public Sub About()
Attribute About.VB_Description = "Adoedc OCX V1.0 By Mario Flores"
Attribute About.VB_UserMemId = -552
    Abouts.Show 1 'Displayed only at Design Time in Propertys
End Sub
