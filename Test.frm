VERSION 5.00
Object = "*\ASQLListview.vbp"
Begin VB.Form frmTest 
   Caption         =   "SQL ListView Demo"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin Sql_ListView.SQLListView SQLListView1 
      Height          =   2415
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Column Header"
      Height          =   330
      Left            =   3045
      TabIndex        =   5
      Top             =   3150
      Width           =   1275
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "Value"
      Height          =   330
      Left            =   2100
      TabIndex        =   4
      Top             =   3150
      Width           =   855
   End
   Begin VB.CommandButton cmdListCount 
      Caption         =   "ListCount"
      Height          =   330
      Left            =   1050
      TabIndex        =   3
      Top             =   3150
      Width           =   960
   End
   Begin VB.CommandButton cmdListIndex 
      Caption         =   "Listindex"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   3150
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate"
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database


Private Sub SQLListView_Change()
    MsgBox "change"
End Sub

Private Sub SQLListView_Click()
    MsgBox "click"
End Sub

Private Sub cmdListCount_Click()
    MsgBox SQLListView1.ListCount
End Sub

Private Sub cmdListIndex_Click()
    MsgBox SQLListView1.ListIndex
End Sub

Private Sub cmdValue_Click()
    Dim n As Integer
    n = InputBox("What Column?")
    MsgBox SQLListView1.Value(n)
End Sub

Private Sub Command1_Click()
    Call Requery_slvCurrency
    Exit Sub
    On Error GoTo ERR_HANDLER

    SQLListView1.Clear
    strSQL = "SELECT fldMoney FROM tblTestx"
    Set rs = db.OpenRecordset(strSQL)
    SQLListView1.rs = rs
    SQLListView1.Requery
ERR_HANDLER:
    MsgBox "This error is handled " & Err.Description
End Sub

Private Sub Command2_Click()
    SQLListView1.Clear
End Sub

Private Sub Command3_Click()
    Dim n As Integer
    n = InputBox("What Column?")
    MsgBox SQLListView1.ColumnHeader(n)
End Sub



Private Function Requery_slvCurrency() As Boolean
 Dim strSQL As String
 Dim rs As Recordset

    strSQL = "SELECT * from tblTest"
    Set rs = db.OpenRecordset(strSQL)
    SQLListView1.rs = rs

    SQLListView1.Requery
    'Set rs = db_billing.OpenRecordset(strSQL)
    'Call pub_Openset(strSQL)
    Requery_slvCurrency = True
End Function



Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\test.mdb", False, True)
End Sub









































