VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "FlexGridBinder Tester © 2005 Richard Mewett"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      Height          =   3225
      Left            =   6570
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "FlexGridBinder.frx":0000
      Top             =   1020
      Width           =   3615
   End
   Begin VB.TextBox txtNotes 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   5580
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5070
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CheckBox chkEditOnMouseClick 
      Caption         =   "Edit on Mouse Click"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   7290
      Width           =   4455
   End
   Begin VB.CheckBox chkFullRowSelect 
      Caption         =   "Full Row Select"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   7530
      Value           =   1  'Checked
      Width           =   2025
   End
   Begin VB.ComboBox cboEdit 
      Height          =   315
      Left            =   3030
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2370
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cboDataType 
      Height          =   315
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2220
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   1830
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid fgDatabaseTable 
      Height          =   3255
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   5741
      _Version        =   393216
      RowHeightMin    =   315
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.PictureBox picHeading 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   10215
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FlexGridBinder Class"
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
         Left            =   1260
         TabIndex        =   2
         Top             =   60
         Width           =   1755
      End
      Begin VB.Label lblPurpose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution to add simple edit functionality to an MSFlexGrid Control"
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
         Left            =   1260
         TabIndex        =   3
         Top             =   300
         Width           =   6450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgContacts 
      Height          =   2535
      Left            =   60
      TabIndex        =   10
      Top             =   4710
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   4471
      _Version        =   393216
      RowHeightMin    =   315
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contacts Grid"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   4410
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database Table Grid"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   750
      Width           =   1470
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#############################################################################################################################
'Demo of FlexGridBinder Class
'
'This Class is used to add simple edit functionality to a MSFlexGrid Control.

'I created it when I had to add a grid edit mode to a system that was
'built using MSFlexGrids and it worked for me so I hope someone else may
'find it useful!

'The Edit functionality can be controlled in Code if required via three simple Methods
'.CancelEdit    Cancel a pending Edit
'.EditCell      Start an Edit
'.UpdateCell    Commit a pending Edit

'NOTE:
'The MSFlexGrid has limited Event functionality to assist in adding tightly
'coupled Edit mode, so there are some compromises that could be overcome with
'lower level integration (such as subclassing). I didn't need that for the
'programs I used it with - my aim was to keep the class lightweight
'#############################################################################################################################

Option Explicit

Private Const ROW_SPACE         As Long = 45

Private Const COL_DATATYPE      As Long = 2
Private Const COL_SIZE          As Long = 3

'*************************************************************
'* Look at the mDatabaseTable Events for operational details *
'*************************************************************

Private WithEvents mContacts As cFlexGridBinder
Attribute mContacts.VB_VarHelpID = -1
Private WithEvents mDatabaseTable As cFlexGridBinder
Attribute mDatabaseTable.VB_VarHelpID = -1

Private Function GetDataTypeSize(sDataTypeText As String) As Integer
    Select Case sDataTypeText
    Case "Date"
        GetDataTypeSize = 8
    Case "Double"
        GetDataTypeSize = 8
    Case "Single"
        GetDataTypeSize = 4
    Case "Currency"
        GetDataTypeSize = 8
    Case "Long"
        GetDataTypeSize = 4
    Case "Integer"
        GetDataTypeSize = 2
    End Select
End Function

Private Sub chkEditOnMouseClick_Click()
    With mContacts
        If chkEditOnMouseClick.Value Then
            .EditOnMouse = emeClick
        Else
            .EditOnMouse = emeNone
        End If
    End With
End Sub

Private Sub chkFullRowSelect_Click()
    With mContacts
        .FullRowSelect = chkFullRowSelect.Value
    End With
End Sub

Private Sub Form_Load()
    With cboDataType
        .AddItem "String"
        .AddItem "Date"
        .AddItem "Double"
        .AddItem "Single"
        .AddItem "Currency"
        .AddItem "Long"
        .AddItem "Integer"
    End With

    '#############################################################################################################################
    'Setup the Database Table example
    
    With fgDatabaseTable
        .FormatString = "|<Column|<DataType|<Size|<Nullable"
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        
        .Rows = 1
        .AddItem vbTab & "OrderID" & vbTab & "Long" & vbTab & "4" & vbTab & "No"
        .AddItem vbTab & "OrderRef" & vbTab & "String" & vbTab & "8" & vbTab & "Yes"
        .AddItem vbTab & "OrderDate" & vbTab & "Date" & vbTab & "8" & vbTab & "Yes"
        .AddItem vbTab & "Description" & vbTab & "String" & vbTab & "64" & vbTab & "Yes"
        .AddItem vbTab & "Goods" & vbTab & "Currency" & vbTab & "8" & vbTab & "Yes"
    End With
    
    Set mDatabaseTable = New cFlexGridBinder
    With mDatabaseTable
        Set .Grid = fgDatabaseTable
        
        .Editable = True
        .EditOnMouse = emeClick
        
        'bSelectText is used to select all the Text in an Edit control when Edit starts
        .BindColumn 1, txtEdit, , , , , True
        .BindColumn 2, cboDataType
        
        'sFilter is used to specify the characters the bound control will accept
        'nColorMode=cmGridCell makes the Edit Control inherit the BackColor & ForeColor from the Grid cell
        .BindColumn 3, txtEdit, "0123456789"
        'sListItems can be used to specify a list of entries for the bound ComboxBox
        .BindColumn 4, cboEdit, , , , "Yes|No"

    End With
    
    '#############################################################################################################################
    'Setup the Contacts example
    
    With fgContacts
        .FormatString = "|<Forename|<Surname|<email address|<Key Colour|<Notes"
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .ColWidth(4) = 1000
        .ColWidth(5) = 3750
        .Rows = 1
        
        .AddItem vbTab & "Jon" & vbTab & "Smith" & vbTab & "jsmith@demo.com" & vbTab & "Yellow"
        .AddItem vbTab & "Chris" & vbTab & "Davis" & vbTab & "chris@demo.com" & vbTab & "Green"
        .AddItem vbTab & "Matt" & vbTab & "Johnson" & vbTab & "mj@demo.com" & vbTab & "Blue"
        .AddItem vbTab & "Claire" & vbTab & "Jones" & vbTab & "cjones@demo.com" & vbTab & "Red"
    End With
    
    Set mContacts = New cFlexGridBinder
    With mContacts
        Set .Grid = fgContacts
        
        .Editable = True
        .EditOnMouse = emeNone
        
        'We cannot use the flexSelectByRow mode on the FlexGrid because it stops us from
        'being able to determine columns individually so the FullRowSelect property is
        'used to emulate it by simply painting a Highlighted row
        .FullRowSelect = True
        
        .BindColumn 1, txtEdit
        .BindColumn 2, txtEdit
        .BindColumn 3, txtEdit
        .BindColumn 4, cboEdit, , , , "Black|Blue|Green|Red|White|Yellow"
        .BindColumn 5, txtNotes
    End With
    
    frmTotals.Show vbModeless
End Sub
Private Sub mDatabaseTable_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '#############################################################################################################################
    'This Event is fired after a ValidateEdit Event is accepted
    '#############################################################################################################################
    
    With fgDatabaseTable
        If Col = COL_DATATYPE Then
            .TextMatrix(Row, COL_SIZE) = GetDataTypeSize(.TextMatrix(Row, COL_DATATYPE))
        End If
    End With
End Sub

Private Sub mDatabaseTable_EditRowChanged(ByVal Row As Long)
    '#############################################################################################################################
    'This Event is fired when a Row is changed during an Edit
    '#############################################################################################################################
End Sub


Private Sub mDatabaseTable_MaxCharsReached(ByVal Row As Long, ByVal Col As Long)
    '#############################################################################################################################
    'This Event is fired when a TextBox entry reaches the MaxLength set in the BindColumn method
    '#############################################################################################################################
End Sub

Private Sub mDatabaseTable_MoveControl(EditID As Long, Left As Single, Top As Single, Width As Single, Height As Single)
    '#############################################################################################################################
    'This Event is fired before an Edit Control is made visible. It allows the size and position of the Control
    'to be adjusted from the defaults calculated by the Class.
    
    'The EditID can be used to identity an Edit and is set in the RequestEdit Event
    '#############################################################################################################################
End Sub


Private Sub mDatabaseTable_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean, EditID As Long)
    '#############################################################################################################################
    'This Event is fired before starting an Edit. You can set Cancel to True to stop
    'the edit from starting. The EditID is used to allow you to identify an Edit within
    'the MoveControl Event (see that Event for description).
    '#############################################################################################################################
    
    'Can we edit the Size?
    If Col = COL_SIZE Then
        With fgDatabaseTable
            'We can only change the size of "String" columns - so prevent Edit of other DataTypes
            If .TextMatrix(Row, COL_DATATYPE) <> "String" Then
                Cancel = True
            End If
        End With
    End If
End Sub


Private Sub mDatabaseTable_RequestRowChange(ByVal NewRow As Long, Cancel As Boolean)
    '#############################################################################################################################
    'This Event is fired before the current Row is changed from a KeyDown action.
    'Setting Cancel to True will stop the Row changing
    '#############################################################################################################################
End Sub

Private Sub mDatabaseTable_RowPainted(ByVal Row As Long)
    '#############################################################################################################################
    'This Event is fired after a Row is formatted. This happens if the FullRowSelect property is True
    '#############################################################################################################################
End Sub

Private Sub mDatabaseTable_ValidateEdit(ByVal Row As Long, ByVal Col As Long, NewData As Variant, Cancel As Boolean)
    '#############################################################################################################################
    'This Event is fired to validate an Edit before it is accepted. Set Cancel to True
    'to reject the Edit
    '#############################################################################################################################
End Sub


Private Sub txtNotes_Change()
    Dim dHeight As Single
    Dim dMaxHeight As Single
    
    dHeight = mContacts.GetCellHeight(txtNotes.Text, 5) + ROW_SPACE
    
    With fgContacts
        dMaxHeight = (.Height - .RowHeightMin) - ROW_SPACE
        
        If dHeight < .RowHeightMin Then
            dHeight = .RowHeightMin
        ElseIf dHeight > dMaxHeight Then
            dHeight = dMaxHeight
        End If
            
        If .RowHeight(.Row) <> dHeight Then
            .RowHeight(.Row) = dHeight
            txtNotes.Height = .RowHeight(.Row) - ROW_SPACE
            
            If txtNotes.Top + txtNotes.Height > .Height Then
                mContacts.LockEvents = True
                
                If .TopRow < .Rows - 1 Then
                    .TopRow = .TopRow + 1
                End If
                
                txtNotes.Top = .RowPos(.Row) + .Top + (.GridLineWidth * Screen.TwipsPerPixelY * 4)
            
                mContacts.LockEvents = False
            End If
        End If
    End With

End Sub


