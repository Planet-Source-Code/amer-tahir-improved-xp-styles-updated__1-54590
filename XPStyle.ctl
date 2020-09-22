VERSION 5.00
Begin VB.UserControl XPStyle 
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "XPStyle.ctx":0000
   Picture         =   "XPStyle.ctx":22DA
   ScaleHeight     =   810
   ScaleWidth      =   810
   ToolboxBitmap   =   "XPStyle.ctx":45B4
End
Attribute VB_Name = "XPStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ____________________________________________________________________________________________________________
'|                                                                                                            |
'|                    In the name of Allah, the Merciful, the Compassionate.                                  |
'|                                ~~~~~~~~~~~~~~~~~~~~~~~~~~~                                                 |
'|                                                                                                            |
'|                                          XPStyle                                                           |
'|                                       Version 1.00                                                         |
'|                                                                                                            |
'|                              * - This module was written by:                                               |
'|                              -------------------------------                                               |
'|                                         Amer Tahir                                                         |
'|                                        amer@amer.cc                                                        |
'|                                                                                                            |
'|                  If you have any questions, feedback, thoughts or anything to share..                      |
'|                                  Please e-mail me immediately! :D                                          |
'|                                                                                                            |
'|____________________________________________________________________________________________________________|

'===================================================
'
' In this new and improved version, you don't
' need any external manifest XML file. This technique
' is used by all Microsoft's apps. The manifest XML
' is embedded in the app as a resource.
' This XP Styling is really easy to implement. Just
' add the XPStyle user control and manifest.res
' resource file in your VB project and place the
' user control on the first form which loads in your
' app. Well, if you like, you can place the user
' control on all forms as well (it doesn't make any problem)
' Compile your app, run it and VOILA! Its XP Styled!
'
'===================================================

Option Explicit

'API which initializes XP Styles on common controls
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200

'These APIs prevent shutdown crashes
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long

'This is the main sub where the XP Styles are initialized
Private Sub UserControl_Initialize()
    Dim iccex As tagInitCommonControlsEx
    iccex.lngSize = LenB(iccex)
    iccex.lngICC = ICC_USEREX_CLASSES
    InitCommonControlsEx iccex
    
    'this is to prevent crash
    m_hMod = LoadLibrary("shell32.dll")
End Sub

'We want our control to remain sized to the picture
Private Sub UserControl_Resize()
    UserControl.Width = 810
    UserControl.Height = 810
End Sub

Private Sub UserControl_Terminate()
    'this is to prevent crash
    FreeLibrary m_hMod
End Sub
