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
' <VB WATCH>
Const VBWMODULE = "frmDeDup"
' </VB WATCH>

Private Sub cboTables_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>


2      Dim colFields As Collection
3      Dim icount As Integer

4      lstFields.Clear

5      Set colFields = New Collection

6      Set colFields = FieldNames(filename, cboTables.List(cboTables.ListIndex))
7      For icount = 1 To colFields.Count
8          lstFields.AddItem colFields(icount)
9      Next icount

10     Set colFields = Nothing

11     chkSelect.Value = False

' <VB WATCH>
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cboTables_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSelect_Click
' DateTime  : 11/07/02 11:40
' Author    : bushells
' Purpose   : Resets listview ticks depending on value of Select All check box
'---------------------------------------------------------------------------------------
'
Private Sub chkSelect_Click()
' <VB WATCH>
13         On Error GoTo vbwErrHandler
' </VB WATCH>
14     Dim icount As Integer

15     If lstFields.ListCount = 0 Then
16         chkSelect.Value = False
17         Exit Sub
18     End If

19     If chkSelect.Value Then

20         For icount = 0 To lstFields.ListCount - 1
21            lstFields.Selected(icount) = True
22         Next icount
23     Else
24         For icount = 0 To lstFields.ListCount - 1
25            lstFields.Selected(icount) = False
26         Next icount

27     End If
28     lstFields.Refresh
' <VB WATCH>
29         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkSelect_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
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
' <VB WATCH>
30         On Error GoTo vbwErrHandler
' </VB WATCH>

31     Dim dProgress As Double
32     Dim icount As Integer
33     Dim boolFieldChosen As Boolean

       'Obviously check if a database has been chosen
34     If LenB(lblDatabase.Caption) = 0 Then
35         MsgBox "A database has not been chosen"
36         Exit Sub
37     End If

38     boolFieldChosen = False

       'Then check to see if any of the listitems (fields) have been chosen to dedup on
39     For icount = 0 To lstFields.ListCount - 1
40         If lstFields.Selected(icount) = True Then
41             boolFieldChosen = True
42             Exit For
43         End If
44     Next icount

45     If Not boolFieldChosen Then
46         MsgBox "No fields have been chosen to De-Duplicate on", vbExclamation
47         Exit Sub
48     End If


49     FormatConnection
50     GetConnection

51     dProgress = GetRecordCount

       'Check if we have got any records to dedup
52     If dProgress = 0 Then
53         MsgBox "No records on database"
54         Exit Sub
55     End If
56     MousePointer = vbHourglass

       'Setup progress bar values
57     pbRecords.Max = dProgress
58     pbRecords.Min = 0
59     pbRecords.Value = 0

60     CreateTables

61     adorsOriginalTable.MoveFirst

62     LoadArray
63     LoadNextArray
64     pbRecords.Value = iProgress
65     adorsOriginalTable.MoveNext

66     Do Until adorsOriginalTable.EOF
67         DeDupRecords
68         pbRecords.Value = iProgress
69         adorsOriginalTable.MoveNext
70         lbldel.Caption = CStr(iCodeSame)
71         DoEvents
72     Loop

73     MousePointer = vbDefault
74     lbldel.Caption = CStr(iCodeSame)

75     Set oConnection = Nothing
76     iProgress = 0

77     iRecRead = 0
78     iCodeSame = 0
' <VB WATCH>
79         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDeDup_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdExit_Click
' DateTime  : 11/07/02 12:18
' Author    : bushells
' Purpose   : Only close connection if it was originally set
'---------------------------------------------------------------------------------------
'
Private Sub cmdExit_Click()
' <VB WATCH>
80         On Error GoTo vbwErrHandler
' </VB WATCH>

81     If boolDatabaseChosen Then
82         Set oConnection = Nothing
83     End If

84     End

' <VB WATCH>
85         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdExit_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub dlbDir_Change()
' <VB WATCH>
86         On Error GoTo vbwErrHandler
' </VB WATCH>
87     flbFile.Path = dlbDir.Path
' <VB WATCH>
88         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "dlbDir_Change"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub dlbDrive_Change()
' <VB WATCH>
89         On Error GoTo vbwErrHandler
' </VB WATCH>
90     On Error GoTo errortrap
91     dlbDir.Path = Left$(dlbDrive.Drive, 1) & ":\"
92     errortrap:
93     If Err.Number = 68 Then
94         MsgBox "Device not ready or unavailable", vbExclamation
95     End If
' <VB WATCH>
96         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "dlbDrive_Change"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub flbFile_Click()
' <VB WATCH>
97         On Error GoTo vbwErrHandler
' </VB WATCH>

98     Dim colTables As Collection
99     Dim icount As Integer

100    MousePointer = vbHourglass
101    cboTables.Clear

102    Set colTables = New Collection

103    filename = flbFile.Path
104    If Right$(filename, 1) <> "\" Then
105         filename = filename & "\"
106    End If
107    filename = filename & flbFile.filename
108    lblDatabase.Caption = filename

109    Set colTables = NonSystemTables(filename)
110    If colTables Is Nothing Then
111        MsgBox "Unable to access the database"
112        Set colTables = Nothing
113        MousePointer = vbDefault
114        Exit Sub
115    End If

116    For icount = 1 To colTables.Count
117        cboTables.AddItem colTables(icount)
118    Next icount

119    Set colTables = Nothing

120    cboTables.ListIndex = 0


121    MousePointer = vbDefault

' <VB WATCH>
122        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "flbFile_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 11/07/02 12:17
' Author    : bushells
' Purpose   : Set filter for Access Databases
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
' <VB WATCH>
123        On Error GoTo vbwErrHandler
' </VB WATCH>
124    flbFile.Pattern = "*.mdb"
' <VB WATCH>
125        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Load"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub
Private Sub mnuChooseDatabase_Click()
' <VB WATCH>
126        On Error GoTo vbwErrHandler
' </VB WATCH>
127    dlbDrive.SetFocus
' <VB WATCH>
128        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuChooseDatabase_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub mnuDatabaseDedup_Click()
' <VB WATCH>
129        On Error GoTo vbwErrHandler
' </VB WATCH>
130    cmdDeDup_Click
' <VB WATCH>
131        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuDatabaseDedup_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub mnuDatabaseFields_Click()
' <VB WATCH>
132        On Error GoTo vbwErrHandler
' </VB WATCH>
133     lstFields.SetFocus
' <VB WATCH>
134        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuDatabaseFields_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub mnuDatabaseSelectAll_Click()
' <VB WATCH>
135        On Error GoTo vbwErrHandler
' </VB WATCH>
136    If chkSelect.Value Then
137     chkSelect.Value = 0
138    Else
139        chkSelect.Value = 1
140    End If

' <VB WATCH>
141        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuDatabaseSelectAll_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub mnuDatabaseTables_Click()
' <VB WATCH>
142        On Error GoTo vbwErrHandler
' </VB WATCH>
143     cboTables.SetFocus
' <VB WATCH>
144        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuDatabaseTables_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

Private Sub mnuFileExit_Click()
' <VB WATCH>
145        On Error GoTo vbwErrHandler
' </VB WATCH>
146    cmdExit_Click
' <VB WATCH>
147        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mnuFileExit_Click"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
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
' <VB WATCH>
148        On Error GoTo vbwErrHandler
' </VB WATCH>

149    If arrHoldDupData(0).FieldToCheck <> adorsOriginalTable(arrFieldPositions(0)) Then
150        LoadArray
151        LoadNextArray
152    Else
153        LoadArray
154        If Not CheckDiff Then
               'Duplicate Record
155            iCodeSame = iCodeSame + 1
156            UpdateDeletedTable
157            adorsOriginalTable.Delete
158        End If
159    End If


' <VB WATCH>
160        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DeDupRecords"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckDiff
' DateTime  : 11/07/02 12:28
' Author    : bushells
' Purpose   : Checks each dedup field to see if its changed
'---------------------------------------------------------------------------------------
'
Private Function CheckDiff() As Boolean
' <VB WATCH>
161        On Error GoTo vbwErrHandler
' </VB WATCH>

162    Dim icount As Integer
163    Dim iLoop As Integer

164    icount = UBound(arrDupData())

165    For iLoop = 0 To icount
166    If arrHoldDupData(iLoop).FieldToCheck <> arrDupData(iLoop).FieldToCheck Then
167        CheckDiff = True
168        Exit Function
169    End If
170    Next iLoop

171    CheckDiff = False

' <VB WATCH>
172        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CheckDiff"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnection
' DateTime  : 11/07/02 11:47
' Author    : bushells
' Purpose   : Forms connection to databse and opens it
'---------------------------------------------------------------------------------------
'
Private Sub GetConnection()
' <VB WATCH>
173        On Error GoTo vbwErrHandler
' </VB WATCH>
174    Dim strConnectionInfo As String

175    boolDatabaseChosen = False

176    Set oConnection = New ADODB.Connection

177    With oConnection
178            .CursorLocation = adUseServer
179            .Mode = adModeReadWrite
180    End With

181    strConnectionInfo = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";JET OLEDB:Database Password=" + sPassword
182    oConnection.Open strConnectionInfo

183    Set adorsOriginalTable = New ADODB.Recordset
184    adorsOriginalTable.Open sSqlConnection, oConnection, adOpenStatic, adLockOptimistic

185    boolDatabaseChosen = True

' <VB WATCH>
186        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetConnection"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub
Private Function GetRecordCount() As Double
' <VB WATCH>
187        On Error GoTo vbwErrHandler
' </VB WATCH>

188    Dim dRecordCount As Double
189    dRecordCount = adorsOriginalTable.RecordCount
190    GetRecordCount = dRecordCount

' <VB WATCH>
191        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetRecordCount"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function
'---------------------------------------------------------------------------------------
' Procedure : LoadArray
' DateTime  : 11/07/02 12:19
' Author    : bushells
' Purpose   : Loads an array with the current records fields chosen for deduplication
'---------------------------------------------------------------------------------------
'
Private Sub LoadArray()
' <VB WATCH>
192        On Error GoTo vbwErrHandler
' </VB WATCH>
193    Dim icount As Integer
194    Dim iElements As Integer

       'Set value for progress bar
195    iProgress = iProgress + 1

       'pintFieldCount is the number of fields chosen to depdup on - set in GetSQLFieldNames
196    iElements = pintFieldCount - 1

197    ReDim arrHoldDupData(iElements)

198    For icount = 0 To iElements
199        If IsNull(adorsOriginalTable(arrFieldPositions(icount))) Then
200            arrHoldDupData(icount).FieldToCheck = vbNullString
201        Else
202            arrHoldDupData(icount).FieldToCheck = adorsOriginalTable(arrFieldPositions(icount))
203        End If
204    Next icount


' <VB WATCH>
205        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadArray"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
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
' <VB WATCH>
206        On Error GoTo vbwErrHandler
' </VB WATCH>
207    Dim icount As Integer
208    Dim iElements As Integer

209    iElements = pintFieldCount - 1

210    ReDim arrDupData(iElements)

211    For icount = 0 To iElements
212        If IsNull(adorsOriginalTable(arrFieldPositions(icount))) Then
213            arrDupData(icount).FieldToCheck = vbNullString
214        Else
215            arrDupData(icount).FieldToCheck = adorsOriginalTable(arrFieldPositions(icount))
216        End If
217    Next icount

' <VB WATCH>
218        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadNextArray"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateDeletedTable
' DateTime  : 11/07/02 12:27
' Author    : bushells
' Purpose   : Inserts into the deleted table the record we are about to delete
'---------------------------------------------------------------------------------------
'
Private Sub UpdateDeletedTable()
' <VB WATCH>
219        On Error GoTo vbwErrHandler
' </VB WATCH>
220    Dim adoUpdateStat As New ADODB.Recordset
221    Dim I As Integer
222    Dim sSqlUpdate As String
223    sSqlUpdate = "SELECT * FROM [Deleted" + pstrTableName + "]"

224    adoUpdateStat.Open sSqlUpdate, oConnection, adOpenStatic, adLockOptimistic

225    adoUpdateStat.AddNew

226    With adorsOriginalTable
227        For I = 0 To .Fields.Count - 1
228        adoUpdateStat.Fields(I).Value = .Fields(I).Value
229        Next
230    End With

231    adoUpdateStat.Update
232    Set adoUpdateStat = Nothing

' <VB WATCH>
233        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "UpdateDeletedTable"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
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
' <VB WATCH>
234        On Error GoTo vbwErrHandler
' </VB WATCH>

235    Dim adoDropTables As New ADODB.Recordset
236    Dim adoCreateTables As New ADODB.Recordset
237    Dim ssqlCreation As String

238    On Error Resume Next

239    ssqlCreation = "DROP TABLE [Deleted" + pstrTableName + "]"

240    adoDropTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic

241    On Error GoTo vbwErrHandler

242    ssqlCreation = "SELECT [" + pstrTableName + "].* INTO [Deleted" + pstrTableName + "] FROM [" + pstrTableName + "];"
243    adoCreateTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic
244    ssqlCreation = "DELETE * FROM [Deleted" + pstrTableName + "]"
245    adoDropTables.Open ssqlCreation, oConnection, adOpenStatic, adLockOptimistic

246    Set adoDropTables = Nothing
247    Set adoCreateTables = Nothing

' <VB WATCH>
248        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CreateTables"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NonSystemTables
' DateTime  : 11/07/02 11:45
' Author    : bushells
' Purpose   : Get the names of the tables in the chosen database
'---------------------------------------------------------------------------------------
'
Public Function NonSystemTables(dbPath As String) As Collection
' <VB WATCH>
249        On Error GoTo vbwErrHandler
' </VB WATCH>

250    Dim td As DAO.TableDef
251    Dim db As DAO.Database
252    Dim colTables As Collection
253    Dim sFormatPassword As String


254    sPassword = Trim$(txtPassword.Text)

255    If sPassword <> "" Then
256        sFormatPassword = "MS ACCESS;PWD=" + sPassword
257        Set db = Workspaces(0).OpenDatabase(dbPath, False, False, sFormatPassword)
258    Else
259        Set db = Workspaces(0).OpenDatabase(dbPath)
260    End If

261    Set colTables = New Collection

262     For Each td In db.TableDefs

263        If td.Attributes >= 0 And td.Attributes <> dbHiddenObject _
                       And td.Attributes <> 2 Then

264              colTables.Add td.Name
265        End If
266      Next
267    db.Close
268    Set NonSystemTables = colTables

' <VB WATCH>
269        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NonSystemTables"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function

'---------------------------------------------------------------------------------------
' Procedure : FieldNames
' DateTime  : 11/07/02 11:44
' Author    : bushells
' Purpose   : Get the field names contained in the table chosen
'---------------------------------------------------------------------------------------
'
Private Function FieldNames(dbPath As String, TableName As String) As Collection
' <VB WATCH>
270        On Error GoTo vbwErrHandler
' </VB WATCH>

271    Dim oCol As Collection
272    Dim db As DAO.Database
273    Dim oTD As DAO.TableDef
274    Dim lCount As Long, lCtr As Long
275    Dim sFormatPassword As String

276    sPassword = Trim$(txtPassword.Text)

277    If sPassword <> "" Then
278        sFormatPassword = "MS ACCESS;PWD=" + sPassword
279        Set db = Workspaces(0).OpenDatabase(dbPath, False, False, sFormatPassword)
280    Else
281        Set db = Workspaces(0).OpenDatabase(dbPath)
282    End If

283    Set oTD = db.TableDefs(TableName)
284    Set oCol = New Collection
285    With oTD
286        lCount = .Fields.Count
287          For lCtr = 0 To lCount - 1
288            oCol.Add .Fields(lCtr).Name
289        Next
290    End With

291    db.Close
292    Set FieldNames = oCol

' <VB WATCH>
293        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FieldNames"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Function
'---------------------------------------------------------------------------------------
' Procedure : FormatConnection
' DateTime  : 11/07/02 11:42
' Author    : bushells
' Purpose   : Builds SQL String
'---------------------------------------------------------------------------------------
'
Private Sub FormatConnection()
' <VB WATCH>
294        On Error GoTo vbwErrHandler
' </VB WATCH>

295    pstrTableName = cboTables.List(cboTables.ListIndex)
296    sSqlConnection = vbNullString
297    sSqlConnection = sSqlConnection + "SELECT *"
298    sSqlConnection = sSqlConnection + " FROM ["
299    sSqlConnection = sSqlConnection + pstrTableName
300    sSqlConnection = sSqlConnection + "] ORDER BY "
301    GetSQLFieldNames

' <VB WATCH>
302        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FormatConnection"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
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
' <VB WATCH>
303        On Error GoTo vbwErrHandler
' </VB WATCH>

304    Dim icount As Integer
305    Dim iFieldCount As Integer
306    Dim boolFirst As Boolean

307    boolFirst = True

308    For icount = 0 To lstFields.ListCount - 1
309        If lstFields.Selected(icount) = True Then
310            ReDim Preserve arrFieldPositions(iFieldCount)
311            arrFieldPositions(iFieldCount) = icount
312            iFieldCount = iFieldCount + 1
313            If Not boolFirst Then
314                sSqlConnection = sSqlConnection + ", "
315                boolFirst = False
316            End If
317            sSqlConnection = sSqlConnection + "[" + pstrTableName + "].[" + lstFields.List(icount) + "]"
318            boolFirst = False
319        End If
320    Next icount

       'Set a varibale to the number of fields chosen, Used in LoadArray and LoadNextArray
321    pintFieldCount = iFieldCount

' <VB WATCH>
322        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetSQLFieldNames"
    Select Case MsgBox("Error " & Err.Number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

' </VB WATCH>
End Sub


' <VB WATCH> <VBWATCHFINALPROC>
' Procedure added by VB Watch
Private Sub Form_Initialize()
    vbwInitialize ' Initialize VB Watch
End Sub
' </VB WATCH>
