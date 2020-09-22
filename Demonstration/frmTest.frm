VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test XPStyle"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin XPStyle_Demo.XPStyle XPStyle1 
      Left            =   4500
      Top             =   2115
      _extentx        =   1429
      _extenty        =   1429
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1545
      LargeChange     =   50
      Left            =   4995
      Max             =   100
      TabIndex        =   11
      Top             =   1710
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   225
      Max             =   100
      TabIndex        =   10
      Top             =   3420
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1740
      Left            =   225
      TabIndex        =   2
      Top             =   1575
      Width           =   4575
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   1425
         Left            =   120
         ScaleHeight     =   1425
         ScaleWidth      =   4335
         TabIndex        =   3
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton cmdNewWin 
            Caption         =   "New Window"
            Height          =   375
            Left            =   1485
            TabIndex        =   13
            Top             =   0
            Width           =   1410
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmTest.frx":0000
            Left            =   1395
            List            =   "frmTest.frx":000D
            TabIndex        =   12
            Text            =   "Combo Box"
            Top             =   495
            Width           =   2805
         End
         Begin VB.TextBox txtSample 
            Height          =   375
            Left            =   2985
            TabIndex        =   9
            Text            =   "empty"
            Top             =   0
            Width           =   1230
         End
         Begin VB.CommandButton cmdMsgBox 
            Caption         =   "MsgBox"
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option 1"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option 2"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option 3"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Checkbox"
            Height          =   255
            Left            =   1350
            TabIndex        =   4
            Top             =   900
            Width           =   1200
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frmTest.frx":0026
      Height          =   645
      Left            =   225
      TabIndex        =   1
      Top             =   765
      Width           =   5100
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTest.frx":0101
      Height          =   510
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   5100
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMsgBox_Click()
    MsgBox "Pretty Cool huh?"
End Sub

Private Sub cmdNewWin_Click()
    frmTest2.Show
End Sub
