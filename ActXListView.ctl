VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SQLListView 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "ActXListView.ctx":0000
   ScaleHeight     =   1470
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ActXListView.ctx":06CA
   Begin MSComctlLib.ListView LV 
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1402
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "SQLListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'///Changes - 8/30/98 hid the propety Listindex from property page

Private m_strColumn     As String
Private m_intListindex  As Integer
Private m_bolEnabled    As Boolean
Private m_strCompareVal As String
Private m_Rs            As Object   '   This is for Recordset


'Event Declarations:
Event Change()
Event Click()
Event dblClick()
Event Keydown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'************* PROPERTIES **************************************
Public Property Get Value(Optional intIndex As Integer = 0) As Variant
'    Get the Value for the control which is the text in
'    the indicated column (intIndex).  If not column is
'    passed then the first/left column is the default
'    Parameters: column index
'    Returns: control value, NULL = nothing selected

    ' Any error results in a NULL value
    On Error Resume Next
    If intIndex = 0 Then
        If LV.SelectedItem.Selected = False Then
            m_value = Null
        Else
            m_value = LV.SelectedItem
        End If
    Else
        If LV.SelectedItem.Selected = False Then
            m_value = Null
        Else
            m_value = LV.SelectedItem.SubItems(intIndex)
        End If
    End If
    If Err <> 0 Then m_value = Null
    Value = m_value
End Property

Public Property Get ColumnHeader(intIndex As Integer) As String
'    Get the Text for the columnheader indicated by the
'    index parameter
'    Parameters: column index
'    Returns: Text for the indicated columnheader
    On Error GoTo ERR_HANDLER
    m_value = LV.ColumnHeaders(intIndex + 1).Text
    ColumnHeader = m_value
    Exit Property
ERR_HANDLER:
    Call Object_Err_Handler
End Property

Public Property Get ColumnHeaderCount() As Integer
'    Get the number of Columnheaders/Query Fields
'    Parameters: none
'    Returns: number of Columnheaders/Query Fields
On Error GoTo ERR_HANDLER
    m_ColumnHeaderCount = LV.ColumnHeaders.Count
    ColumnHeaderCount = m_ColumnHeaderCount
    Exit Property
ERR_HANDLER:
    Call Object_Err_Handler
End Property

Public Property Get Listindex() As Integer
'    Get the index of the selected listitem
'    Parameters: none
'    Returns: index of the selected listitem
    On Error Resume Next
    
    ' This is a test for design mode - if sql is empty then
    ' the app. can't be running.  Otherwise, this sub triggers
    ' an error in design mode.  Note ActiveX controls are "live"
    ' in BOTH design and runtime.
    If m_Rs Is Nothing Then Exit Property
    
    
    If LV.SelectedItem.Selected = False Then
        m_intListindex = 0
    Else
        m_intListindex = LV.SelectedItem.Index
    End If
    Listindex = m_intListindex
    If Err <> 0 Then m_intListindex = 0
End Property

Public Property Let Listindex(New_ListIndex As Integer)
'    Sets the listindex property and highlights the corresponding
'    listitem
'    Parameters: listindex value
'    Returns: N/A
    ' For 0 set all items to selected = false
On Error GoTo ERR_HANDLER
    If New_ListIndex = 0 Then
        For C = 1 To LV.ListItems.Count
            ' get the selected item and unselect it
            If LV.ListItems(C).Selected = True Then
                LV.ListItems(C).Selected = False
                Exit For
            End If
        Next
        m_strCompareVal = "" 'set Change check = ""
    Else
        LV.ListItems(New_ListIndex).Selected = True
    End If
    m_intListindex = New_ListIndex
    PropertyChanged "Listindex"
    'Change only for a non 0 listindex
    If New_ListIndex <> 0 Then RaiseEvent Change
    Exit Property
ERR_HANDLER:
    Call Object_Err_Handler
End Property


Public Property Let Rs(new_Rs As Object)
    Set m_Rs = new_Rs
End Property

Public Property Get Rs() As Object
    Rs = m_Rs
End Property


Public Property Let Enabled(newEnabled As Boolean)
'    Sets the Enabled property to control to True or False
'    Parameters: True to Enable, False to Disable control
'    Returns: N/A
    m_bolEnabled = newEnabled
    LV.Enabled = m_bolEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
'    Gets the Enabled property of control (True or False)
'    Parameters: None
'    Returns: True - control is enabled, False - control is disabled
    Enabled = m_bolEnabled
End Property

Private Sub LV_BeforeLabelEdit(Cancel As Integer)
'    Overrides Defaul ListView functionality that allows
'    Users to type into the listitems 1st column
'    and overwrite the text
    Cancel = True
End Sub

'************  EVENTS ***************************

Private Sub LV_Click()
    RaiseEvent Click
End Sub

Private Sub LV_DblClick()
    RaiseEvent dblClick
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' compare previous LV.selecteditem value (m_strCompareVal)
    ' to current LV.selecteditem - if different raise the CHANGE
    ' event. Note - [CHANGE] will be raised BEFORE [CLICK] if a
    ' [change] has occurred after the [click].  Otherwise just [click]
    ' will be raised.
    If m_strCompareVal <> LV.SelectedItem Then
        RaiseEvent Change
        m_strCompareVal = LV.SelectedItem
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ERR_HANDLER
    RaiseEvent MouseDown(Button, Shift, x, y)
    ' Prevents error that would occurr on last line
    If LV.ListItems.Count = 0 Then Exit Sub
    Exit Sub
ERR_HANDLER:
    Call Object_Err_Handler
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER
    RaiseEvent Keydown(KeyCode, Shift)
    Exit Sub
ERR_HANDLER:
    Call Object_Err_Handler
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    'Default state for this control is ENABLED
    m_bolEnabled = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", m_bolEnabled)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_bolEnabled)
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo ERR_HANDLER
    LV.SortKey = ColumnHeader.Index - 1
    ' Sort Ascending/Descending
    If m_strColumn = ColumnHeader Then
        If LV.SortOrder = lvwDescending Then
            LV.SortOrder = lvwAscending
        Else
            LV.SortOrder = lvwDescending
        End If
    End If
    m_strColumn = ColumnHeader
    LV.Sorted = True
    Exit Sub
ERR_HANDLER:
    Call Object_Err_Handler
End Sub

'***************** USER CONTROL EVENTS ******************



Private Sub UserControl_Resize()
'    Allows Constituant ListView contol to allways FILL the
'    UserControl dimensions in Design mode.  Gives appearance
'    that developer is actually resizing ListView control.
    On Error GoTo ERR_HANDLER
    LV.Height = Text1.Parent.Height
    LV.Width = Text1.Parent.Width
    Exit Sub
ERR_HANDLER:
    Call Object_Err_Handler
End Sub


'********* METHODS *****************************************
Public Sub Clear()
    LV.ListItems.Clear
End Sub

Public Sub Requery(Optional ClearHeaders As Boolean = False)
 Dim intTotCount As Integer
 Dim intCount1 As Integer, intCount2 As Integer
 Dim colNew As ColumnHeader, NewLine As ListItem


    On Error GoTo ERR_HANDLER

    ' Clear the ListView control.
    LV.ListItems.Clear


    If m_Rs Is Nothing Then Exit Sub
    
    ' Reset columnheaders if necessary
    If ClearHeaders Or LV.ColumnHeaders.Count = 0 Then
        LV.ColumnHeaders.Clear
        For intCount1 = 0 To m_Rs.Fields.Count - 1
            Set colNew = LV.ColumnHeaders.Add(, , m_Rs(intCount1).Name)
        Next intCount1
    End If
  
    LV.View = 3    ' Set View property to 'Report'.
    If m_Rs.RecordCount = 0 Then Exit Sub
    
    ' Set Total Records Counter.
    m_Rs.MoveLast
    intTotCount = m_Rs.RecordCount
    m_Rs.MoveFirst

    ' Loop through recordset and add Items to the control.
    For intCount1 = 1 To intTotCount
        '///[changed] 8/25/98 to allow for formatted numbers (I.e. $5,000)
        '[changed] 8/25/98 If IsNumeric(rs(0).Value) Then
        '[changed] 8/25/98     Set NewLine = LV.ListItems.Add(, , LTrim(RTrim(str(rs(0).Value))))
        '[changed] 8/25/98 Else
            '[changed] 8/30/98 Took off LTRIM to allow for sorting of numbers
            Set NewLine = LV.ListItems.Add(, , pub_SQL2Text(m_Rs(0)))
        '[changed] 8/25/98 End If

        For intCount2 = 1 To m_Rs.Fields.Count - 1
                NewLine.SubItems(intCount2) = pub_SQL2Text(m_Rs(intCount2), True)
        Next intCount2
        m_Rs.MoveNext
    Next intCount1
    
    m_strCompareVal = ""  'reset value used to raise Change event
      
    LV.LabelWrap = False
    On Error Resume Next '//  this nex LOC will trip if form isn't loaded
                         '// (I.e. the control is requeried in the form_load event
    If LV.Enabled = True Then LV.SetFocus
    Exit Sub
ERR_HANDLER:
    ' Ignore Error 94 which indicates you passed a NULL value.
    If Err = 94 Then
        Resume Next
    Else
        Call Object_Err_Handler
    End If
End Sub

Public Function SetColumnWidth(intIndex As Integer, ByVal New_Width As Integer)
'    Sets the columnwidth of a specific listview column
'    Parameters:  Column to change, New specified width
'    Returns: Number of listitems in the control
'    Based on Base Level 0
    On Error GoTo ERR_HANDLER
    If IsMissing(intIndex) Then intIndex = 1
    LV.ColumnHeaders(intIndex + 1).Width = New_Width
    Exit Function
ERR_HANDLER:
    Call Object_Err_Handler
End Function

Public Function Listcount() As Integer
'    Gets the number of listitems in the control
'    Parameters:  None
'    Returns: Number of listitems in the control
    Listcount = LV.ListItems.Count
End Function

'********* PRIVATE FUNCTIONS/SUBS *********************

Private Function pub_SQL2Text(val As Variant, _
                Optional bolRtrim As Boolean = True) As Variant
'    Converts recordset field into a string
'    Parameters:  val = recordset field retrieved from SQL Database
'    Returns: string
    If IsNull(val) Then
        pub_SQL2Text = ""
    ElseIf IsDate(val) Then
        pub_SQL2Text = CDate(val)
    Else
        If bolRtrim Then
            pub_SQL2Text = RTrim(val)   'Take off trailing spaces
        Else
            pub_SQL2Text = CStr(val)
        End If
    End If
End Function

Private Function Object_Err_Handler()
    Err.Raise Err.Number, , Err.Description
End Function


