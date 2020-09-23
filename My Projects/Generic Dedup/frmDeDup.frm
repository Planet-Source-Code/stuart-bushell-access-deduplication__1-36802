VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDeDup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access De Duplication"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9585
   ControlBox      =   0   'False
   Icon            =   "frmDeDup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Height          =   720
      Left            =   75
      TabIndex        =   16
      Top             =   3915
      Width           =   2640
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   390
         Left            =   1320
         TabIndex        =   18
         Top             =   225
         Width           =   930
      End
      Begin VB.CommandButton cmdDeDup 
         Caption         =   "&DeDup"
         Height          =   390
         Left            =   285
         TabIndex        =   17
         Top             =   225
         Width           =   930
      End
   End
   Begin VB.Frame fraFields 
      Caption         =   "F&ields"
      Height          =   3765
      Left            =   5895
      TabIndex        =   9
      Top             =   1065
      Width           =   3585
      Begin VB.CheckBox chkSelect 
         Caption         =   "&Select All"
         Height          =   225
         Left            =   180
         TabIndex        =   15
         Top             =   315
         Width           =   3195
      End
      Begin VB.ListBox lstFields 
         Height          =   2985
         Left            =   135
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   630
         Width           =   3300
      End
   End
   Begin VB.Frame fraTables 
      Caption         =   "&Tables"
      Height          =   990
      Left            =   5910
      TabIndex        =   8
      Top             =   60
      Width           =   3585
      Begin VB.ComboBox cboTables 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   375
         Width           =   3330
      End
   End
   Begin MSComctlLib.ProgressBar pbRecords 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   4980
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame fraDataBase 
      Caption         =   "&Choose Database"
      Height          =   3300
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   5745
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3015
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2760
         Width           =   2340
      End
      Begin VB.DriveListBox dlbDrive 
         Height          =   315
         Left            =   225
         TabIndex        =   0
         Top             =   375
         Width           =   2655
      End
      Begin VB.DirListBox dlbDir 
         Height          =   1665
         Left            =   210
         TabIndex        =   1
         Top             =   930
         Width           =   2670
      End
      Begin VB.FileListBox flbFile 
         Height          =   2235
         Left            =   2985
         TabIndex        =   3
         Top             =   360
         Width           =   2400
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Database Password (if required)"
         Height          =   300
         Left            =   210
         TabIndex        =   14
         Top             =   2820
         Width           =   2640
      End
   End
   Begin VB.Label lblDatabase 
      Height          =   300
      Left            =   75
      TabIndex        =   13
      Top             =   3465
      Width           =   5745
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress"
      Height          =   285
      Left            =   15
      TabIndex        =   12
      Top             =   4725
      Width           =   780
   End
   Begin VB.Label lblMisMatched 
      Caption         =   "Duplicate Records Found :"
      Height          =   240
      Left            =   2880
      TabIndex        =   11
      Top             =   4380
      Width           =   1935
   End
   Begin VB.Label lbldel 
      Height          =   255
      Left            =   4860
      TabIndex        =   10
      Top             =   4380
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "Data&base"
      Begin VB.Menu mnuDatabaseDedup 
         Caption         =   "&DeDup"
      End
      Begin VB.Menu mnuChooseDatabase 
         Caption         =   "&Choose Database"
      End
      Begin VB.Menu mnuDatabaseTables 
         Caption         =   "&Tables"
      End
      Begin VB.Menu mnuDatabaseFields 
         Caption         =   "F&ields"
         Begin VB.Menu mnuDatabaseSelectAll 
            Caption         =   "&Select All"
         End
      End
   End
End
Attribute VB_Name = "frmDeDup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDeDup
' DateTime  : 11/07/02 11:39
' Author    : bushells
' Purpose   :'This project allows a de-duplication process on an Access Database which allows
'               more than 10 fields to be selected (which is the limit using the Access Wizard).
'               It will create a table in the same database which will contain all records
'               deleted from the table selected for de-duplication.
'---------------------------------------------------------------------------------------
Option Explicit
Private adorsOriginalTable As ADODB.Recordset
Private oConnection As ADODB.Connection

Private boolDatabaseChosen As Boolean

Private iProgress As Integer
Private iRecRead As Integer
Private pintFieldCount As Integer
Private iCodeSame As Integer

Private sSqlConnection As String
Private pstrTableName As String
Private filename As String
Private sPassword As String

Private Type DupData
    FieldToCheck As Variant
End Type

Private arrHoldDupData() As DupData
Private arrDupData() As DupData
Private arrFieldPositions() As Integer

'---------------------------------------------------------------------------------------
' Procedure : cboTables_Click
' DateTime  : 11/07/02 11:38
' Author    : bushells
' Purpose   : This process fills a listbox with all the fields from
'               the table selected from the combo box.
'---------------------------------------------------------------------------------------
'
Private Sub cboTables_Click()


Dim colFields As Collection
Dim icount As Integer

lstFields.Clear

Set colFields = New Collection

Set colFields = FieldNames(filename, cboTables.List(cboTables.ListIndex))
For icount = 1 To colFields.Count
    lstFields.AddItem colFields(icount)
Next icount

Set colFields = Nothing

chkSelect.Value = False
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSelect_Click
' DateTime  : 11/07/02 11:40
' Author    : bushells
' Purpose   : Resets listview ticks depending on value of Select All check box
'---------------------------------------------------------------------------------------
'
Private Sub chkSelect_Click()
Dim icount As Integer

If lstFields.ListCount = 0 Then
    chkSelect.Value = False
    Exit Sub
End If

If chkSelect.Value Then

    For icount = 0 To lstFields.ListCount - 1
       lstFields.Selected(icount) = True
    Next icount
Else
    For icount = 0 To lstFields.ListCount - 1
       lstFields.Selected(icount) = False
    Next icount

End If
lstFields.Refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdDeDup_Click
' DateTime  : 11/07/02 11:41
' Author    : bushells
' Purpose   :'This process controls the looping and initilaisation of the
'               deduplication process
'---------------------------------------------------------------------------------------
'
Private Sub cmdDeDup_Click()

Dim dProgress As Double
Dim icount As Integer
Dim boolFieldChosen As Boolean

'Obviously check if a database has been chosen
If LenB(lblDatabase.Caption) = 0 Then
    MsgBox "A database has not been chosen"
    Exit Sub
End If

boolFieldChosen = False

'Then check to see if any of the listitems (fields) have been chosen to dedup on
For icount = 0 To lstFields.ListCount - 1
    If lstFields.Selected(icount) = True Then
        boolFieldChosen = True
        Exit For
    End If
Next icount

If Not boolFieldChosen Then
    MsgBox "No fields have been chosen to De-Duplicate on", vbExclamation
    Exit Sub
End If


FormatConnection
GetConnection

dProgress = GetRecordCount

'Check if we have got any records to dedup
If dProgress = 0 Then
    MsgBox "No records on database"
    Exit Sub
End If
MousePointer = vbHourglass

'Setup progress bar values
pbRecords.Max = dProgress
pbRecords.Min = 0
pbRecords.Value = 0

CreateTables

adorsOriginalTable.MoveFirst

LoadArray
LoadNextArray
pbRecords.Value = iProgress
adorsOriginalTable.MoveNext

Do Until adorsOriginalTable.EOF
    DeDupRecords
    pbRecords.Value = iProgress
    adorsOriginalTable.MoveNext
    lbldel.Caption = CStr(iCodeSame)
    DoEvents
Loop

MousePointer = vbDefault
lbldel.Caption = CStr(iCodeSame)

Set oConnection = Nothing
iProgress = 0

iRecRead = 0
iCodeSame = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdExit_Click
' DateTime  : 11/07/02 12:18
' Author    : bushells
' Purpose   : Only close connection if it was originally set
'---------------------------------------------------------------------------------------
'
Private Sub cmdExit_Click()

If boolDatabaseChosen Then
    Set oConnection = Nothing
End If

End

End Sub

Private Sub dlbDir_Change()
flbFile.Path = dlbDir.Path
End Sub

Private Sub dlbDrive_Change()
On Error GoTo errortrap
dlbDir.Path = Left$(dlbDrive.Drive, 1) & ":\"
errortrap:
If Err.Number = 68 Then
    MsgBox "Device not ready or unavailable", vbExclamation
End If
End Sub

Private Sub flbFile_Click()

Dim colTables As Collection
Dim icount As Integer

MousePointer = vbHourglass
cboTables.Clear

Set colTables = New Collection

filename = flbFile.Path
If Right$(filename, 1) <> "\" Then
     filename = filename & "\"
End If
filename = filename & flbFile.filename
lblDatabase.Caption = filename

Set colTables = NonSystemTables(filename)
If colTables Is Nothing Then
    MsgBox "Unable to access the database"
    Set colTables = Nothing
    MousePointer = vbDefault
    Exit Sub
End If

For icount = 1 To colTables.Count
    cboTables.AddItem colTables(icount)
Next icount

Set colTables = Nothing

cboTables.ListIndex = 0


MousePointer = vbDefault
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 11/07/02 12:17
' Author    : bushells
' Purpose   : Set filter for Access Databases
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
flbFile.Pattern = "*.mdb"
End Sub
Private Sub mnuChooseDatabase_Click()
dlbDrive.SetFocus
End Sub

Private Sub mnuDatabaseDedup_Click()
cmdDeDup_Click
End Sub

Private Sub mnuDatabaseFields_Click()
 lstFields.SetFocus
End Sub

Private Sub mnuDatabaseSelectAll_Click()
If chkSelect.Value Then
 chkSelect.Value = 0
Else
    chkSelect.Value = 1
End If
  
End Sub

Private Sub mnuDatabaseTables_Click()
 cboTables.SetFocus
End Sub

Private Sub mnuFileExit_Click()
cmdExit_Click
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DeDupRecords
' DateTime  : 11/07/02 12:29
' Author    : bushells
' Purpose   : If the first dedup field has changed then we have a new record, otherwise
'               check to see if the rest of the fileds are the same (call Checkdiff)
'---------------------------------------------------------------------------------------
'
Private Sub DeDupRecords()

If arrHoldDupData(0).FieldToCheck <> adorsOriginalTable(arrFieldPositions(0)) Then
    LoadArray
    LoadNextArray
Else
    LoadArray
    If Not CheckDiff Then
        'Duplicate Record
        iCodeSame = iCodeSame + 1
        UpdateDeletedTable
        adorsOriginalTable.Delete
    End If
End If

  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckDiff
' DateTime  : 11/07/02 12:28
' Author    : bushells
' Purpose   : Checks each dedup field to see if its changed
'---------------------------------------------------------------------------------------
'
Private Function CheckDiff() As Boolean

Dim icount As Integer
Dim iLoop As Integer

icount = UBound(arrDupData())

For iLoop = 0 To icount
If arrHoldDupData(iLoop).FieldToCheck <> arrDupData(iLoop).FieldToCheck Then
    CheckDiff = True
    Exit Function
End If
Next iLoop

CheckDiff = False
  
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnection
' DateTime  : 11/07/02 11:47
' Author    : bushells
' Purpose   : Forms connection to databse and opens it
'---------------------------------------------------------------------------------------
'
Private Sub GetConnection()
Dim strConnectionInfo As String

boolDatabaseChosen = False

Set oConnection = New ADODB.Connection

With oConnection
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
End With

strConnectionInfo = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";JET OLEDB:Database Password=" + sPassword
oConnection.Open strConnectionInfo

Set adorsOriginalTable = New ADODB.Recordset
adorsOriginalTable.Open sSqlConnection, oConnection, adOpenStatic, adLockOptimistic

boolDatabaseChosen = True
  
End Sub
Private Function GetRecordCount() As Double

Dim dRecordCount As Double
dRecordCount = adorsOriginalTable.RecordCount
GetRecordCount = dRecordCount
  
End Function
'---------------------------------------------------------------------------------------
' Procedure : LoadArray
' DateTime  : 11/07/02 12:19
' Author    : bushells
' Purpose   : Loads an array with the current records fields chosen for deduplication
'---------------------------------------------------------------------------------------
'
Private Sub LoadArray()
Dim icount As Integer
Dim iElements As Integer

'Set value for progress bar
iProgress = iProgress + 1

'pintFieldCount is the number of fields chosen to depdup on - set in GetSQLFieldNames
iElements = pintFieldCount - 1

ReDim arrHoldDupData(iElements)

For icount = 0 To iElements
    If IsNull(adorsOriginalTable(arrFieldPositions(icount))) Then
        arrHoldDupData(icount).FieldToCheck = vbNullString
    Else
        arrHoldDupData(icount).FieldToCheck = adorsOriginalTable(arrFieldPositions(icount))
    End If
Next icount

  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : LoadNextArray
' DateTime  : 11/07/02 12:26
' Author    : bushells
' Purpose   : Loads an array with the next records fields chosen for deduplication for
'               later comparison
'---------------------------------------------------------------------------------------
'
Private Sub LoadNextArray()
Dim icount As Integer
Dim iElements As Integer

iElements = pintFieldCount - 1

ReDim arrDupData(iElements)

For icount = 0 To iElements
    If IsNull(adorsOriginalTable(arrFieldPositions(icount))) Then
        arrDupData(icount).FieldToCheck = vbNullString
    Else
        arrDupData(icount).FieldToCheck = adorsOriginalTable(arrFieldPositions(icount))
    End If
Next icount
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateDeletedTable
' DateTime  : 11/07/02 12:27
' Author    : bushells
' Purpose   : Inserts into the deleted table the record we are about to delete
'---------------------------------------------------------------------------------------
'
Private Sub UpdateDeletedTable()
Dim adoUpdateStat As New ADODB.Recordset
Dim I As Integer
Dim sSqlUpdate As String
sSqlUpdate = "SELECT * FROM [Deleted" + pstrTableName + "]"

adoUpdateStat.Open sSqlUpdate, oConnection, adOpenStatic, adLockOptimistic

adoUpdateStat.AddNew

With adorsOriginalTable
    For I = 0 To .Fields.Count - 1
    adoUpdateStat.Fields(I).Value = .Fields(I).Value
    Next
End With

adoUpdateStat.Update
Set adoUpdateStat = Nothing
  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CreateTables
' DateTime  : 11/07/02 11:30
' Author    : bushells
' Purpose   : To create a table based on the one we are to dedup to hold those
'               records that we will delete (just in case!)
'---------------------------------------------------------------------------------------
'
Private Sub CreateTables()

Dim adoDropTables As New ADODB.Recordset
Dim adoCreateTables As New ADODB.Recordset
Dim ssqlCreation As String

On Error Resume Next

ssqlCreation = "DROP TABLE [Deleted" + pstrTableName + "]"

adoDropTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic

On Error GoTo 0

ssqlCreation = "SELECT [" + pstrTableName + "].* INTO [Deleted" + pstrTableName + "] FROM [" + pstrTableName + "];"
adoCreateTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic
ssqlCreation = "DELETE * FROM [Deleted" + pstrTableName + "]"
adoDropTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic

Set adoDropTables = Nothing
Set adoCreateTables = Nothing
  
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NonSystemTables
' DateTime  : 11/07/02 11:45
' Author    : bushells
' Purpose   : Get the names of the tables in the chosen database
'---------------------------------------------------------------------------------------
'
Public Function NonSystemTables(dbPath As String) As Collection

Dim td As DAO.TableDef
Dim db As DAO.Database
Dim colTables As Collection
Dim sFormatPassword As String


sPassword = Trim$(txtPassword.Text)

If sPassword <> "" Then
    sFormatPassword = "MS ACCESS;PWD=" + sPassword
    Set db = Workspaces(0).OpenDatabase(dbPath, False, False, sFormatPassword)
Else
    Set db = Workspaces(0).OpenDatabase(dbPath)
End If

Set colTables = New Collection

 For Each td In db.TableDefs

    If td.Attributes >= 0 And td.Attributes <> dbHiddenObject _
                And td.Attributes <> 2 Then

          colTables.Add td.Name
    End If
  Next
db.Close
Set NonSystemTables = colTables
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : FieldNames
' DateTime  : 11/07/02 11:44
' Author    : bushells
' Purpose   : Get the field names contained in the table chosen
'---------------------------------------------------------------------------------------
'
Private Function FieldNames(dbPath As String, TableName As String) As Collection

Dim oCol As Collection
Dim db As DAO.Database
Dim oTD As DAO.TableDef
Dim lCount As Long, lCtr As Long
Dim sFormatPassword As String

sPassword = Trim$(txtPassword.Text)

If sPassword <> "" Then
    sFormatPassword = "MS ACCESS;PWD=" + sPassword
    Set db = Workspaces(0).OpenDatabase(dbPath, False, False, sFormatPassword)
Else
    Set db = Workspaces(0).OpenDatabase(dbPath)
End If

Set oTD = db.TableDefs(TableName)
Set oCol = New Collection
With oTD
    lCount = .Fields.Count
      For lCtr = 0 To lCount - 1
        oCol.Add .Fields(lCtr).Name
    Next
End With

db.Close
Set FieldNames = oCol

End Function
'---------------------------------------------------------------------------------------
' Procedure : FormatConnection
' DateTime  : 11/07/02 11:42
' Author    : bushells
' Purpose   : Builds SQL String
'---------------------------------------------------------------------------------------
'
Private Sub FormatConnection()

pstrTableName = cboTables.List(cboTables.ListIndex)
sSqlConnection = vbNullString
sSqlConnection = sSqlConnection + "SELECT *"
sSqlConnection = sSqlConnection + " FROM ["
sSqlConnection = sSqlConnection + pstrTableName
sSqlConnection = sSqlConnection + "] ORDER BY "
GetSQLFieldNames
  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : GetSQLFieldNames
' DateTime  : 11/07/02 11:33
' Author    : bushells
' Purpose   : Builds the part of the SQL string to get the data ordered by the
'               fields chosen to dedup on. Sets an array to hold the field positions
'               of the fields chosen to dedup on (arrFieldPositions).
'---------------------------------------------------------------------------------------
'
Private Sub GetSQLFieldNames()

Dim icount As Integer
Dim iFieldCount As Integer
Dim boolFirst As Boolean

boolFirst = True

For icount = 0 To lstFields.ListCount - 1
    If lstFields.Selected(icount) = True Then
        ReDim Preserve arrFieldPositions(iFieldCount)
        arrFieldPositions(iFieldCount) = icount
        iFieldCount = iFieldCount + 1
        If Not boolFirst Then
            sSqlConnection = sSqlConnection + ", "
            boolFirst = False
        End If
        sSqlConnection = sSqlConnection + "[" + pstrTableName + "].[" + lstFields.List(icount) + "]"
        boolFirst = False
    End If
Next icount

'Set a varibale to the number of fields chosen, Used in LoadArray and LoadNextArray
pintFieldCount = iFieldCount
  
End Sub

