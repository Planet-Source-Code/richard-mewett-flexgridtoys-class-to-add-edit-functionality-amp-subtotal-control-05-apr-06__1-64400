VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FlexGridSubtotal Tester © 2006 Richard Mewett"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      Height          =   1065
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Totals.frx":0000
      Top             =   720
      Width           =   7935
   End
   Begin VB.PictureBox picHeading 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8070
      TabIndex        =   2
      Top             =   0
      Width           =   8130
      Begin VB.Label lblPurpose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution to add simple Subtotals to an MSFlexGrid Control"
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
         TabIndex        =   5
         Top             =   300
         Width           =   5835
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FlexGridSubtotal Control"
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
         TabIndex        =   4
         Top             =   60
         Width           =   2070
      End
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
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
   End
   Begin FlexGridBinder.RMFlexGridSubtotal fgsData 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   5430
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   3555
      Left            =   90
      TabIndex        =   0
      Top             =   1860
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6271
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With fgData
        .FormatString = "|<Paint|>Quantity|>Price|<In Stock|>Delivery|<Notes|<Order No|<Reference|<Store Code"
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 750
        .ColWidth(5) = 1000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        
        .Rows = 1
        .AddItem vbTab & "Black" & vbTab & "6" & vbTab & "10.00" & vbTab & "Yes" & vbTab & "3.00"
        .AddItem vbTab & "Blue" & vbTab & "8" & vbTab & "12.50" & vbTab & "Yes" & vbTab & "3.00"
        .AddItem vbTab & "Gold" & vbTab & "1" & vbTab & "16.50" & vbTab & "No" & vbTab & "3.00"
        .AddItem vbTab & "Red" & vbTab & "8" & vbTab & "12.50" & vbTab & "Yes" & vbTab & "3.00"
        .AddItem vbTab & "Yellow" & vbTab & "3" & vbTab & "12.50" & vbTab & "Yes" & vbTab & "3.00"
        .AddItem vbTab & "White" & vbTab & "20" & vbTab & "9.50" & vbTab & "Yes" & vbTab & "6.00"
    End With
    
    With fgsData
        Set .Grid = fgData
        .ForeColor = vbRed
        .Font.Bold = True
        
        .BindColumn 2
        .BindColumn 3, "#.00"
        .BindColumn 5, "#.00"
        .Recalculate
    End With
End Sub


