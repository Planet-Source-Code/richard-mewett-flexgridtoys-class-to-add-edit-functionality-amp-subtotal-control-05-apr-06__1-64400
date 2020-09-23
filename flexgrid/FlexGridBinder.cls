VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlexGridBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################################################
'Title:     FlexGridBinder (Add simple Edit functionality to an MSFlexGrid)
'Author:    Richard Mewett
'Created:   05/03/05
'Version:   1.0.3 (25th February 2006)

'Copyright © 2005 Richard Mewett. All rights reserved.
'
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "FlexGridBinder". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.


'*********************************************************************
'Notes:
'To use this class successfuly the FlexGrid should be set to:
'HighLight=flexHighlightNever
'SelectionMode=flexSelectionFree
'*********************************************************************

'History: (dd/mm/yy)
'25/02/06   Replaced global SelTextOnEdit property with SelectText option on BindColumn
'           Added ColorMode option on BindColumn

'##############################################################################################################
'Windows API
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'##############################################################################################################
Private Const DEF_EDITBACKCOLOR As Long = &HC0FFFF

Public Enum ColorModeEnum
    cmDefault = 0
    cmGridCell = 1
    cmControl = 2
End Enum

Public Enum EditOnMouseEnum
    emeNone = 0
    emeClick = 1
    emeDblClick = 2
End Enum

Private Type ColSettingsType
    Items As Collection
    Ctrl As Control
    nMaxLength As Integer
    sFilter As String
    sFormat As String
    bSelectText As Boolean
    nColorMode As ColorModeEnum
End Type

#If False Then
    Private cmDefault, cmGridCell, cmControl
    Private emeNone, emeClick, emeDblClick
    Private Row, Col
#End If

Private WithEvents mGrid As MSFlexGrid
Attribute mGrid.VB_VarHelpID = -1
Private WithEvents mComboBox As VB.ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mTextBox As VB.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private mGeneric As VB.Control

Private mOriginalHighLight As Integer
Private mOriginalSelectionMode As Integer
Private mEditID As Long

Private mCols() As ColSettingsType
Private mEditBackcolor As Long
Private mEditForecolor As Long
Private mEditCol As Long
Private mEditRow As Long
Private mHighlightedRow As Long
Private mEditOnMouse As EditOnMouseEnum
Private mChangeTabStops As Boolean
Private mDataChanged As Boolean
Private mEditable As Boolean
Private mEditPending As Boolean
Private mFullRowSelect As Boolean
Private mLockEvents As Boolean

Private mAllowUserResizing As Integer

Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event EditRowChanged(ByVal Row As Long)
Public Event RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean, EditID As Long)
Public Event MoveControl(EditID As Long, Left As Single, Top As Single, Width As Single, Height As Single)
Public Event RequestRowChange(ByVal NewRow As Long, Cancel As Boolean)
Public Event RowPainted(ByVal Row As Long)
Public Event ValidateEdit(ByVal Row As Long, ByVal Col As Long, NewData As Variant, Cancel As Boolean)

Public Event MaxCharsReached(ByVal Row As Long, ByVal Col As Long)

Public Sub BindColumn(ByVal nColumn As Integer, Ctrl As Control, Optional ByVal sFilter As String, Optional ByVal nMaxLength As Integer, Optional ByVal sFormat As String, Optional ByVal sListItems As String, Optional bSelectText As Boolean, Optional nColorMode As ColorModeEnum = cmDefault)
    '#############################################################################################################################
    'This is used to set the Edit Control used on a specific column.
    
    'Ctrl       - The control used for editing that Column (a single Control can be used for any number of columns).
    'Note:      The Class uses WithEvents declaration for a TextBox or a ComboBox respectively, which allows it to
    '           trap Events
    
    '           Any other Control uses a standard "As Control" declaration so cannot process Events. I did not want
    '           to add controls that would not always be availble in a standard VB project (such as a DateTimePicker).
    '           For these controls:
    '           1.) Put your own Code in the Events of these controls
    '           2.) You could add more Control declarations to the Class if really required.
    '               i.e Private WithEvents mDateTimePicker As DateTimePicker
    
    'sFilter    - Specify the characters allowed in the KeyPress event of the edit Control (ignore for no filtering)
    'nMaxLength - Used to raise MaxCharsReached Event when you need to restrict length of input
    'sFormat    - Format mask passed to Format$() function after Edit
    'sListItems - List if Items for a ComboBox, seperated by "|". eg. "One|Two|Three|Four"
    'bSelectText- Select the text when starting an Edit
    'nColorMode - Set the method used to determine Control BackColor & ForeColor
    '#############################################################################################################################
    
    Dim nCount As Integer
    Dim sText() As String
    
    If UBound(mCols) = 0 Then
        ReDim mCols(mGrid.Cols)
    ElseIf nColumn > UBound(mCols) Then
        ReDim Preserve mCols(nColumn)
    End If
    
    Set mCols(nColumn).Ctrl = Ctrl
    
    With mCols(nColumn)
        .nMaxLength = nMaxLength
        .sFilter = sFilter
        .sFormat = sFormat
        .bSelectText = bSelectText
        .nColorMode = nColorMode
        
        Set .Items = New Collection
        
        If Len(sListItems) > 0 Then
            sText() = Split(sListItems, "|")
            For nCount = LBound(sText) To UBound(sText)
                .Items.Add sText(nCount)
            Next nCount
        End If
    End With
End Sub

Public Sub CancelEdit()
    '#############################################################################################################################
    'Cancel a pending edit - does not commit changes
    '#############################################################################################################################
    
    Dim ctlEdit As Control
    
    If mEditPending Then
        mEditPending = False
            
        Set ctlEdit = mCols(mEditCol).Ctrl
        If TypeOf ctlEdit Is ComboBox Then
            Set mComboBox = Nothing
        ElseIf TypeOf ctlEdit Is TextBox Then
            Set mTextBox = Nothing
        Else
            Set mGeneric = Nothing
        End If
            
        With mGrid
            .Enabled = True
            .SetFocus
            ctlEdit.Visible = False
            
            ProcessTabStops True
            
            .AllowUserResizing = mAllowUserResizing
            .HighLight = mOriginalHighLight
            .SelectionMode = mOriginalSelectionMode
        End With
    End If
End Sub

Public Property Get ChangeTabStops() As Boolean
    '#############################################################################################################################
    'If set to True, the control resets the TabStop property of all Controls contained by its Parent, to
    'allow tab to be trapped within the Grid. (a better technical solution is to use IOleInPlaceActiveObject but
    'this requiures a lot of code and can make the program very unstable in my experience!)
    '#############################################################################################################################
    
    ChangeTabStops = mChangeTabStops
End Property

Public Property Let ChangeTabStops(ByVal NewValue As Boolean)
    mChangeTabStops = NewValue
End Property

Private Sub Class_Initialize()
    ReDim mCols(0)
    
    mEditBackcolor = DEF_EDITBACKCOLOR
    mEditForecolor = vbWindowText
    
    mEditOnMouse = emeDblClick
End Sub

Public Property Get DataChanged() As Boolean
    DataChanged = mDataChanged
End Property

Public Property Let DataChanged(ByVal NewValue As Boolean)
    mDataChanged = NewValue
End Property

Public Property Get Editable() As Boolean
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
End Property

Public Property Get EditBackcolor() As Long
    EditBackcolor = mEditBackcolor
End Property

Public Property Let EditBackcolor(ByVal NewValue As Long)
    mEditBackcolor = NewValue
End Property

Public Function EditCell(ByVal lRow As Long, ByVal lCol As Long) As Boolean
    '#############################################################################################################################
    'Start an Edit. This will process the Control bound to the Column:
    '1. Set Backcolor & Forecolor
    '2. Set Size & Position
    '3. Set the Text
    '4. Make it Visible
    
    'Notice that a RequestEdit Event is raised before the Edit Control is displayed - this allows the
    'program to stop an edit occuring if required
    
    'The MoveControl Event is raised to allow override of the default Size & Position
    '#############################################################################################################################
    
    Dim ctlEdit As Control
    Dim dLeft As Single
    Dim dHeight As Single
    Dim dTop As Single
    Dim dWidth As Single
    Dim nCount As Integer
    Dim bCancel As Boolean
    Dim sText As Variant
    
    If mEditable Then
        If (lCol < 0) Or (lCol >= UBound(mCols)) Then
            Exit Function
        ElseIf (lRow < 1) Or (lRow >= mGrid.Rows) Then
            Exit Function
        End If
        
        Set ctlEdit = mCols(lCol).Ctrl
        
        If Not ctlEdit Is Nothing Then
            UpdateCell
            
            mEditID = 0
            RaiseEvent RequestEdit(lRow, lCol, bCancel, mEditID)
            If bCancel Then
                Exit Function
            End If
            
            EditCell = True
            
            With mGrid
                mAllowUserResizing = .AllowUserResizing
                .AllowUserResizing = flexResizeNone
                
                mOriginalHighLight = .HighLight
                .HighLight = flexHighlightNever
                
                mOriginalSelectionMode = .SelectionMode
                .SelectionMode = flexSelectionFree
                
                mEditCol = lCol
                mEditRow = lRow
                
                If (.Row <> mEditRow) Then
                    .Row = mEditRow
                    If (.RowPos(mEditRow) + .RowHeight(mEditRow)) > .Height Then
                        .TopRow = .TopRow + 1
                    ElseIf (.RowPos(mEditRow) < .RowHeight(mEditRow)) Then
                        If .TopRow > 1 Then
                            .TopRow = .TopRow - 1
                        End If
                    End If
                    
                    RaiseEvent EditRowChanged(mEditRow)
                End If
                
                If TypeOf ctlEdit Is ComboBox Then
                    Set mComboBox = ctlEdit
                    
                    dHeight = ctlEdit.Height
                    dWidth = .ColWidth(mEditCol) - (.GridLineWidth * Screen.TwipsPerPixelX * 2)
                    dLeft = .ColPos(mEditCol) + .Left + (.GridLineWidth * Screen.TwipsPerPixelX * 4)
                    dTop = .RowPos(mEditRow) + .Top + (.GridLineWidth * Screen.TwipsPerPixelY * 2)
                    
                    RaiseEvent MoveControl(mEditID, dLeft, dTop, dWidth, dHeight)
                    
                    ctlEdit.Width = dWidth
                    
                    If mCols(mEditCol).Items.Count > 0 Then
                        ctlEdit.Clear
                        For Each sText In mCols(mEditCol).Items
                            ctlEdit.AddItem sText
                        Next sText
                    End If
                    
                    With ctlEdit
                        For nCount = 0 To .ListCount - 1
                            If .List(nCount) = mGrid.TextMatrix(mEditRow, mEditCol) Then
                                .ListIndex = nCount
                                Exit For
                            End If
                        Next nCount
                    End With
                ElseIf TypeOf ctlEdit Is TextBox Then
                    Set mTextBox = ctlEdit
                    
                    dHeight = .RowHeight(mEditRow) - (.GridLineWidth * Screen.TwipsPerPixelY * 3)
                    dWidth = .ColWidth(mEditCol) - (.GridLineWidth * Screen.TwipsPerPixelX * 6)
                    dLeft = .ColPos(mEditCol) + .Left + (.GridLineWidth * Screen.TwipsPerPixelX * 6)
                    dTop = .RowPos(mEditRow) + .Top + (.GridLineWidth * Screen.TwipsPerPixelY * 4)
                    
                    RaiseEvent MoveControl(mEditID, dLeft, dTop, dWidth, dHeight)
                    
                    ctlEdit.Height = dHeight
                    ctlEdit.Width = dWidth
                    
                    Select Case .ColAlignment(mEditCol)
                    Case flexAlignRightTop, flexAlignRightCenter, flexAlignRightBottom
                        ctlEdit.Alignment = vbRightJustify
                    
                    Case flexAlignCenterTop, flexAlignCenterCenter, flexAlignCenterBottom
                        ctlEdit.Alignment = vbCenter
                    
                    Case Else
                        ctlEdit.Alignment = vbLeftJustify
                    
                    End Select
                    
                    ctlEdit.Text = .TextMatrix(mEditRow, mEditCol)
                Else
                    Set mGeneric = ctlEdit
                    dHeight = .RowHeight(mEditRow) - (.GridLineWidth * Screen.TwipsPerPixelY * 3)
                    dWidth = .ColWidth(mEditCol) - (.GridLineWidth * Screen.TwipsPerPixelX * 6)
                    dLeft = .ColPos(mEditCol) + .Left + (.GridLineWidth * Screen.TwipsPerPixelX * 6)
                    dTop = .RowPos(mEditRow) + .Top + (.GridLineWidth * Screen.TwipsPerPixelY * 4)
                    
                    RaiseEvent MoveControl(mEditID, dLeft, dTop, dWidth, dHeight)
                    
                    ctlEdit.Height = dHeight
                    ctlEdit.Width = dWidth
                    
                    On Local Error Resume Next
                    ctlEdit.Text = .TextMatrix(mEditRow, mEditCol)
                    On Error GoTo 0
                End If
                
                With ctlEdit
                    If mCols(lCol).bSelectText Then
                        .SelStart = 0
                        .SelLength = Len(.Text)
                    End If
                    
                    Select Case mCols(lCol).nColorMode
                        Case cmDefault
                            .BackColor = mEditBackcolor
                            .ForeColor = mEditForecolor
    
                        Case cmGridCell
                            mLockEvents = True
                            If mGrid.Row <> mEditRow Then
                                mGrid.Row = mEditRow
                            End If
                            If mGrid.Col <> mEditCol Then
                                mGrid.Col = mEditCol
                            End If
                            .BackColor = mGrid.CellBackColor
                            .ForeColor = mGrid.CellForeColor
                            mLockEvents = False
    
                    End Select
                    
                    .Left = dLeft
                    .Top = dTop
                    
                    .Visible = True
                    .SetFocus
                End With
                
                If (.SelectionMode = flexSelectionByRow) Then
                    .Row = mEditRow
                    .Col = 0
                    .ColSel = .Cols - 1
                End If
            End With
            
            ProcessTabStops False
            
            mEditPending = True
        End If
    End If
End Function

Public Property Get EditForecolor() As Long
    EditForecolor = mEditForecolor
End Property

Public Property Let EditForecolor(ByVal NewValue As Long)
    mEditForecolor = NewValue
End Property

Public Property Get EditOnMouse() As EditOnMouseEnum
    EditOnMouse = mEditOnMouse
End Property

Public Property Let EditOnMouse(ByVal NewValue As EditOnMouseEnum)
    mEditOnMouse = NewValue
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = mFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal NewValue As Boolean)
    If mFullRowSelect And Not NewValue Then
        UnSelectRow mHighlightedRow
        mHighlightedRow = 0
    End If
    
    mFullRowSelect = NewValue
End Property

Public Function GetCellHeight(sText As String, lCol As Long, Optional dWidth As Single) As Single
    '#############################################################################################################################
    'Calculate the height of a Edit Control based on the text specified. This is used when the grid
    'contains a multline textbox and the row must be resized dynamically on input.
    '#############################################################################################################################
    
    Dim R As RECT
    
    With mGrid
        If dWidth = 0 Then
            R.Right = mGrid.ColWidth(lCol) \ Screen.TwipsPerPixelY
        Else
            R.Right = dWidth \ Screen.TwipsPerPixelY
        End If
        
        Call DrawText(GetDC(.Parent.hWnd), sText, Len(sText), R, DT_CALCRECT Or DT_WORDBREAK)
    End With
    
    GetCellHeight = R.Bottom * Screen.TwipsPerPixelY
End Function

Public Property Get Grid() As MSFlexGrid
    Set Grid = mGrid
End Property

Public Property Set Grid(ByVal NewValue As MSFlexGrid)
    Set mGrid = NewValue
End Property

Public Property Get HighlightedRow() As Long
    HighlightedRow = mHighlightedRow
End Property

Public Property Let HighlightedRow(ByVal Row As Long)
    If mFullRowSelect Then
        mHighlightedRow = Row
    End If
End Property

Private Function KeyFilter(KeyAscii As Integer, sFilter As String)
    '#############################################################################################################################
    'Simple Key Filter to block unwanted characters
    '#############################################################################################################################
    
    Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete
        KeyFilter = KeyAscii
    
    Case Else
        If InStr(sFilter, Chr$(KeyAscii)) > 0 Then
            KeyFilter = KeyAscii
        End If
    
    End Select
End Function

Public Property Let LockEvents(NewValue As Boolean)
    mLockEvents = NewValue
End Property

Private Sub mComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    ProcessKeyDown KeyCode, Shift
End Sub

Private Sub mComboBox_Validate(Cancel As Boolean)
    If Not UpdateCell() Then
        Cancel = True
    End If
End Sub

Private Sub mGrid_DblClick()
    If mEditOnMouse = emeDblClick Then
        With mGrid
            EditCell .MouseRow, .MouseCol
        End With
    End If
End Sub

Private Sub mGrid_GotFocus()
    PaintRow
End Sub

Private Sub mGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    With mGrid
        Select Case KeyCode
        Case vbKeyReturn
            If .SelectionMode = flexSelectionByRow Then
                EditCell .Row, mEditCol
            Else
                EditCell .Row, .Col
            End If
        
        End Select
    End With
End Sub

Private Sub mGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (mEditOnMouse = emeClick) And (Button = vbLeftButton) Then
        With mGrid
            EditCell .MouseRow, .MouseCol
        End With
    End If
End Sub

Private Sub mGrid_RowColChange()
    If Not mLockEvents Then
        UpdateCell
        PaintRow
    End If
End Sub

Private Sub mGrid_Scroll()
    mGrid_RowColChange
End Sub

Private Sub mGrid_SelChange()
    If Not mLockEvents Then
        PaintRow
    End If
End Sub

Private Sub MoveCellLeft(Optional bAllowRowScroll As Boolean)
    '#############################################################################################################################
    'Used when navigating Grid with keyboard
    '#############################################################################################################################
    
    Dim lCol As Long
    
    lCol = mEditCol
    Do
        lCol = lCol - 1
        If lCol < 0 Then
            If bAllowRowScroll And (mEditRow > 1) Then
                lCol = mGrid.Cols
                mEditRow = mEditRow - 1
            Else
                Exit Do
            End If
        ElseIf mGrid.ColWidth(lCol) > 0 Then
            EditCell mEditRow, lCol
            Exit Do
        End If
    Loop
End Sub

Private Sub MoveCellRight(Optional bAllowRowScroll As Boolean)
    '#############################################################################################################################
    'Used when navigating Grid with keyboard
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim bCancel As Boolean
    
    lCol = mEditCol
    Do
        lCol = lCol + 1
        If lCol = mGrid.Cols Then
            If bAllowRowScroll And (mEditRow < mGrid.Rows - 1) Then
                RaiseEvent RequestRowChange(mEditRow + 1, bCancel)
                If Not bCancel Then
                    lCol = 0
                    UpdateCell
                    mEditRow = mEditRow + 1
                End If
            Else
                bCancel = True
            End If
        ElseIf mGrid.ColWidth(lCol) > 0 Then
            If EditCell(mEditRow, lCol) Then
                bCancel = True
            End If
        End If
    Loop Until bCancel
End Sub

Private Sub mTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    ProcessKeyDown KeyCode, Shift
End Sub

Private Sub mTextBox_KeyPress(KeyAscii As Integer)
    Dim nMaxLength As Integer
    
    If mEditPending Then
        nMaxLength = mCols(mEditCol).nMaxLength
        If ((Len(mTextBox.Text) >= nMaxLength)) And (nMaxLength > 0) Then
            RaiseEvent MaxCharsReached(mEditRow, mEditCol)
        End If
        
        If Len(mCols(mEditCol).sFilter) > 0 Then
            KeyAscii = KeyFilter(KeyAscii, mCols(mEditCol).sFilter)
        End If
    End If
End Sub

Private Sub mTextBox_Validate(Cancel As Boolean)
    If Not UpdateCell() Then
        Cancel = True
    End If
End Sub

Private Sub PaintRow(Optional bForce As Boolean)
    '#############################################################################################################################
    'This is used to display a selected state on a Row. [See FullRowSelect Property]
    
    'The FlexGrid has its own property for this: SelectionMode=flexSelectByRow
    'The problem is that when a FlexGrid is set to that mode we cannot process Columns as required
    'by an Editable Grid.
    '#############################################################################################################################
    
    Static bPainting As Boolean
    
    Dim lOriginalCol As Long
    Dim lOriginalRow As Long
    Dim nOriginalFillStyle As Integer
    
    If bPainting Then
        Exit Sub
    ElseIf mFullRowSelect Then
        mLockEvents = True
        bPainting = True
        
        With mGrid
            If mHighlightedRow > (.Rows - 1) Then
                mHighlightedRow = .Rows - 1
            End If
            
            If (.Row <> mHighlightedRow) Or bForce Then
                lOriginalCol = .Col
                lOriginalRow = .Row
                nOriginalFillStyle = .FillStyle
                
                .FillStyle = flexFillRepeat
                
                If mHighlightedRow > 0 Then
                    .Row = mHighlightedRow
                    .Col = .FixedCols
                    .ColSel = .Cols - 1
                    .CellBackColor = vbWindowBackground
                    .CellForeColor = vbWindowText
                    
                    RaiseEvent RowPainted(.Row)
                    
                    .Row = lOriginalRow
                End If
                
                If (.Row >= .FixedRows) And (.Rows > 1) Then
                    .Col = .FixedCols
                    .ColSel = .Cols - 1
                    .CellBackColor = vbHighlight
                    .CellForeColor = vbHighlightText
                    
                    RaiseEvent RowPainted(.Row)
                
                    .Col = lOriginalCol
                    .FillStyle = nOriginalFillStyle
                End If
                
                mHighlightedRow = .Row
            End If
        End With
        
        mLockEvents = False
        bPainting = False
    End If
End Sub

Public Sub ProcessKeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        If Shift = 0 Then
            KeyCode = 0
            UpdateCell
        End If
    
    Case vbKeyLeft
         If (Shift And vbShiftMask) Then
            KeyCode = 0
            MoveCellLeft
         End If
    
    Case vbKeyRight
         If (Shift And vbShiftMask) Then
            KeyCode = 0
            MoveCellRight
         End If
         
    Case vbKeyTab
        If (Shift And vbCtrlMask) And (Shift And vbShiftMask) Then
            'Move Previous Control in Tab Order
        ElseIf (Shift And vbCtrlMask) Then
            'Move Next Control in Tab Order
        ElseIf (Shift And vbShiftMask) Then
            'Move Previous Cell
            KeyCode = 0
            MoveCellLeft True
        Else
            'Move Next Cell
            KeyCode = 0
            MoveCellRight True
        End If
         
    Case vbKeyUp
         If (Shift And vbShiftMask) Then
            KeyCode = 0
            EditCell mEditRow - 1, mEditCol
         End If
    
    Case vbKeyDown
         If (Shift And vbShiftMask) Then
            KeyCode = 0
            EditCell mEditRow + 1, mEditCol
         End If
    
    End Select
End Sub

Private Sub ProcessTabStops(bState As Boolean)
    '#############################################################################################################################
    'Update TabStop property on the Parents Controls - [See ChangeTabStops Property]
    '#############################################################################################################################
    
    Dim frm As Form
    Dim ctl As Control
    
    On Local Error Resume Next
    
    If mChangeTabStops Then
        Set frm = mGrid.Parent
        For Each ctl In frm.Controls
            If TypeOf ctl Is Label Then
            ElseIf bState Then
                ctl.TabStop = (Val(ctl.Tag) = 1)
            Else
                ctl.Tag = Abs(ctl.TabStop)
                ctl.TabStop = False
            End If
        Next ctl
    End If
End Sub

Public Sub Refresh()
    PaintRow True
End Sub

Private Sub UnSelectRow(lRow As Long)
    Dim lOriginalCol As Long
    Dim lOriginalRow As Long
    Dim nOriginalFillStyle As Integer
    
    mLockEvents = True
    
    With mGrid
        lOriginalCol = .Col
        lOriginalRow = .Row
        nOriginalFillStyle = .FillStyle
        
        .FillStyle = flexFillRepeat
        
        If lRow > 0 Then
            .Row = lRow
            .Col = .FixedCols
            .ColSel = .Cols - 1
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowText
        End If
            
        .Col = lOriginalCol
        .Row = lOriginalRow
        .FillStyle = nOriginalFillStyle
    End With
    
    mLockEvents = False
End Sub

Public Function UpdateCell() As Boolean
    '#############################################################################################################################
    'Updates the Grid with the data from the Edit Control.
    
    'Notice that a ValidateEdit Event is raised before the update is committed - this allows the
    'program to reject the Update occuring if required
    '#############################################################################################################################
    
    Dim ctlEdit As Control
    Dim bCancel As Boolean
    Dim sValue As String
    
    UpdateCell = True
    
    If mEditPending Then
        mEditPending = False
            
        With mGrid
            Set ctlEdit = mCols(mEditCol).Ctrl
            
            If .TextMatrix(mEditRow, mEditCol) <> ctlEdit.Text Then
                sValue = ctlEdit.Text
                
                RaiseEvent ValidateEdit(mEditRow, mEditCol, sValue, bCancel)
                If bCancel Then
                    UpdateCell = False
                Else
                    If TypeOf ctlEdit Is ComboBox Then
                        Set mComboBox = Nothing
                    ElseIf TypeOf ctlEdit Is TextBox Then
                        Set mTextBox = Nothing
                    Else
                        Set mGeneric = Nothing
                    End If
                    
                    ctlEdit.Visible = False
                    If Len(mCols(mEditCol).sFormat) > 0 Then
                        sValue = Format$(sValue, mCols(mEditCol).sFormat)
                    End If
                    
                    If .MergeRow(mEditRow) Then
                        .Text = sValue
                    Else
                        .TextMatrix(mEditRow, mEditCol) = sValue
                    End If
                    
                    .Enabled = True
                    .SetFocus
                    
                    mDataChanged = True
    
                    RaiseEvent AfterEdit(mEditRow, mEditCol)
                End If
            Else
                .Enabled = True
                .SetFocus
                ctlEdit.Visible = False
            End If
            
            ProcessTabStops True
            
            .AllowUserResizing = mAllowUserResizing
            .HighLight = mOriginalHighLight
            .SelectionMode = mOriginalSelectionMode
        End With
    End If
End Function
