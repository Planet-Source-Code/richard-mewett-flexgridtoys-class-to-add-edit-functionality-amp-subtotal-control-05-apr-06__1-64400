VERSION 5.00
Begin VB.UserControl RMFlexGridSubtotal 
   AutoRedraw      =   -1  'True
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
End
Attribute VB_Name = "RMFlexGridSubtotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#############################################################################################################################
'Title:     FlexGridSubtotal (Add Subtotal functionality to an MSFlexGrid)
'Author:    Richard Mewett
'Created:   26/02/06
'Version:   1.0.0 (26th February 2006)

'Copyright © 2005 Richard Mewett. All rights reserved.
'
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "FlexGridSubtotal". Use of
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

Option Explicit

'#############################################################################################################################
'API
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const CLR_INVALID = &HFFFF

Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X           As Long
    Y           As Long
End Type

'#############################################################################################################################
Private Const DEF_BACKCOLOR     As Long = vbWindowBackground
Private Const DEF_BACKCOLORBKG  As Long = &H808080
Private Const DEF_BORDERSTYLE   As Integer = 1
Private Const DEF_FORECOLOR     As Long = vbWindowText

Public Enum FGSBorderStyleEnum
    bsNone = 0
    bsFixedSingle = 1
End Enum

Public Enum FGSSubtotalEnum
    stNone = 0
    stAverage = 1
    stMin = 2
    stMax = 3
    stSum = 4
End Enum

#If False Then
    Private bsNone, bsFixedSingle
    Private stNone, stAverage, stMin, stMax, stSum
#End If

Private Type udtSubtotalColumn
    sFormat As String
    lColumn As Long
    lBackColor As Long
    lForeColor As Long
    cMax As Currency
    cMin As Currency
    cTotal As Currency
    nMode As FGSSubtotalEnum
End Type

Private mCols() As udtSubtotalColumn

Private mBackColor As OLE_COLOR
Private mBackColorBkg As OLE_COLOR
Private mBorderStyle As Integer
Private mFont As Font
Private mForeColor As OLE_COLOR

Private WithEvents mGrid As MSFlexGrid
Attribute mGrid.VB_VarHelpID = -1

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    mBackColor = NewValue
    Refresh
    
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = mBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal NewValue As OLE_COLOR)
    mBackColorBkg = NewValue
    With UserControl
        .BackColor = mBackColorBkg
    End With
    Refresh
    
    PropertyChanged "BackColorBkg"
End Property

Public Sub BindColumn(ByVal lColumn As Long, Optional ByVal sFormat As String, Optional nMode As FGSSubtotalEnum = stSum, Optional ByVal lBackColor As Long = -1, Optional ByVal lForeColor As Long = -1)
    '#############################################################################################################################
    'This is used to set the Subtotal used on a specific column.
    
    'sFormat    - Format mask passed to Format$() function after Edit
    '#############################################################################################################################
    
    Dim nCount As Integer
    
    nCount = UBound(mCols) + 1
    ReDim Preserve mCols(nCount)
    
    With mCols(nCount)
        .lColumn = lColumn
        .sFormat = sFormat
        .nMode = nMode
        If lBackColor = -1 Then
            .lBackColor = mBackColor
        Else
            .lBackColor = lBackColor
        End If
        If lForeColor = -1 Then
            .lForeColor = mForeColor
        Else
            .lForeColor = lForeColor
        End If
    End With
End Sub

Public Property Get BorderStyle() As FGSBorderStyleEnum
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As FGSBorderStyleEnum)
    mBorderStyle = NewValue
    With UserControl
        .BorderStyle = mBorderStyle
    End With
     
    PropertyChanged "BorderStyle"
End Property

Public Sub ClearColumns()
    ReDim mCols(0)
End Sub

Private Sub DrawLine(hdc As Long, X1 As Long, Y1 As Long, x2 As Long, Y2 As Long, lColor As Long)
    Dim pt As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, pt
    LineTo hdc, x2, Y2
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Sub

Private Sub DrawRect(hdc As Long, rc As RECT, lColor As Long, bFilled As Boolean)
    Dim lNewBrush As Long
  
    lNewBrush = CreateSolidBrush(lColor)
    
    If bFilled Then
        Call FillRect(hdc, rc, lNewBrush)
    Else
        Call FrameRect(hdc, rc, lNewBrush)
    End If

    Call DeleteObject(lNewBrush)
End Sub

Public Property Get Font() As Font
   Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    
    Set UserControl.Font = mFont
    Refresh
     
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    With UserControl
        .ForeColor = mForeColor
    End With
    Refresh
    
    PropertyChanged "ForeColor"
End Property

Public Property Get Grid() As MSFlexGrid
    Set Grid = mGrid
End Property

Public Property Set Grid(ByVal NewValue As MSFlexGrid)
    Set mGrid = NewValue
End Property

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Private Sub mGrid_Scroll()
    Refresh
End Sub

Public Sub Recalculate(Optional Column As Long = -1)
    '#############################################################################################################################
    'This is used to set the calculate the Subtotals per column.
    
    'Column - Optionally specify a specific Column to Recalculate - otherwise calculates all Columns
    '#############################################################################################################################
   
    Dim lCol As Long
    Dim lColStart As Long
    Dim lColEnd As Long
    Dim lRow As Long
    Dim cValue As Currency
    
    If Column >= 0 Then
        lColStart = Column
        lColEnd = Column
    Else
        lColStart = LBound(mCols)
        lColEnd = UBound(mCols)
    End If
    
    For lCol = lColStart To lColEnd
        mCols(lCol).cMax = 0
        mCols(lCol).cMin = 0
        mCols(lCol).cTotal = 0
    Next lCol
    
    With mGrid
        For lRow = 1 To .Rows - 1
            For lCol = lColStart To lColEnd
                cValue = Val(.TextMatrix(lRow, mCols(lCol).lColumn))
                
                If (cValue < mCols(lCol).cMin) Or (lRow = 1) Then
                    mCols(lCol).cMin = cValue
                End If
                
                If (cValue > mCols(lCol).cMax) Or (lRow = 1) Then
                    mCols(lCol).cMax = cValue
                End If
                
                mCols(lCol).cTotal = mCols(lCol).cTotal + cValue
            Next lCol
        Next lRow
    End With
    
    Refresh
End Sub

Public Sub Refresh()
    '#############################################################################################################################
    'Draw the Subtotals!
    '#############################################################################################################################
    
    Dim R As RECT
    Dim lCol As Long
    Dim lX As Long
    Dim lWidth As Long
    Dim sText As String
    
    If Not mGrid Is Nothing Then
        With UserControl
            .Cls
        
            For lCol = 1 To UBound(mCols)
                If mCols(lCol).nMode <> stNone Then
                    lWidth = .ScaleX(mGrid.ColWidth(mCols(lCol).lColumn), vbTwips, vbPixels)
                    lX = .ScaleX(mGrid.ColPos(mCols(lCol).lColumn), vbTwips, vbPixels)
                    
                    SetRect R, lX, 0, lX + lWidth, UserControl.Height
                    DrawRect .hdc, R, TranslateColor(mBackColor), True
                    DrawLine .hdc, lX, 0, lX, UserControl.Height, TranslateColor(mGrid.GridColor)
                    
                    Select Case mCols(lCol).nMode
                        Case stAverage
                            sText = mCols(lCol).cTotal / mGrid.Rows
                        Case stMin
                            sText = mCols(lCol).cMin
                        Case stMax
                            sText = mCols(lCol).cMax
                        Case stSum
                            sText = mCols(lCol).cTotal
                    End Select
                    
                    If Len(mCols(lCol).sFormat) > 0 Then
                        sText = Format$(sText, mCols(lCol).sFormat)
                    End If
                    
                    lWidth = .ScaleX(mGrid.ColWidth(mCols(lCol).lColumn) - (mGrid.GridLineWidth * Screen.TwipsPerPixelX * 2), vbTwips, vbPixels)
                    SetRect R, lX, 0, lX + lWidth, UserControl.Height
                    
                    Select Case mGrid.ColAlignment(mCols(lCol).lColumn)
                        Case flexAlignCenterBottom, flexAlignCenterCenter, flexAlignCenterTop
                            Call DrawText(.hdc, sText, -1, R, DT_CENTER)
                        Case flexAlignRightBottom, flexAlignRightCenter, flexAlignRightTop
                            Call DrawText(.hdc, sText, -1, R, DT_RIGHT)
                        Case Else
                            Call DrawText(.hdc, sText, -1, R, DT_LEFT)
                    End Select
                            
                End If
            Next lCol
            
            .Refresh
        End With
    End If
End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub UserControl_Initialize()
    ReDim mCols(0)
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font
    
    mBackColor = DEF_BACKCOLOR
    mBorderStyle = DEF_BORDERSTYLE
    mForeColor = DEF_FORECOLOR
    
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
        .ForeColor = mForeColor
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    mBackColor = PropBag.ReadProperty("BackColor", DEF_BACKCOLOR)
    mBackColorBkg = PropBag.ReadProperty("BackColorBkg", DEF_BACKCOLORBKG)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", DEF_BORDERSTYLE)
    mForeColor = PropBag.ReadProperty("ForeColor", DEF_FORECOLOR)
    
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
        .ForeColor = mForeColor
    End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    
    Call PropBag.WriteProperty("BackColor", mBackColor, DEF_BACKCOLOR)
    Call PropBag.WriteProperty("BackColorBkg", mBackColorBkg, DEF_BACKCOLORBKG)
    Call PropBag.WriteProperty("BorderStyle", mBorderStyle, DEF_BORDERSTYLE)
    Call PropBag.WriteProperty("ForeColor", mForeColor, DEF_FORECOLOR)
End Sub
