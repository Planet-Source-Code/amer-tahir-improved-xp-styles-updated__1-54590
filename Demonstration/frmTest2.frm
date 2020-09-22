VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Another form"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   270
      TabIndex        =   1
      Top             =   315
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Me"
      Height          =   465
      Left            =   1755
      TabIndex        =   0
      Top             =   1350
      Width           =   1950
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
