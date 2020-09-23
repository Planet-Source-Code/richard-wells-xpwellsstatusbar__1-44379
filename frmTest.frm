VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00E0F0F0&
   Caption         =   "XP StatusBar Control"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrKey 
      Interval        =   100
      Left            =   1200
      Top             =   480
   End
   Begin StatusBarTest.xpWellsStatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2265
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   873
      BackColor       =   14708792
      ForeColor       =   16777215
      ForeColorDissabled=   9915703
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfPanels  =   6
      MaskColor       =   16711935
      ShowGripper     =   -1  'True
      PWidth1         =   100
      pText1          =   ""
      pTTText1        =   "Hello"
      pEnabled1       =   -1  'True
      PanelPicture1   =   "frmTest.frx":0000
      PWidth2         =   100
      pText2          =   "My Computer"
      pTTText2        =   ""
      pEnabled2       =   -1  'True
      PanelPicture2   =   "frmTest.frx":001C
      PWidth3         =   80
      pText3          =   "Internet"
      pTTText3        =   ""
      pEnabled3       =   -1  'True
      PanelPicture3   =   "frmTest.frx":036E
      PWidth4         =   25
      pText4          =   ""
      pTTText4        =   "Privacy Report"
      pEnabled4       =   -1  'True
      PanelPicture4   =   "frmTest.frx":06C0
      PWidth5         =   25
      pText5          =   ""
      pTTText5        =   "You Have New Mail"
      pEnabled5       =   -1  'True
      PanelPicture5   =   "frmTest.frx":0A12
      PWidth6         =   40
      pText6          =   "CAPS"
      pTTText6        =   ""
      pEnabled6       =   -1  'True
      PanelPicture6   =   "frmTest.frx":0D64
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Keyboard API
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long

Private kbArray As KeyboardBytes
Private Const VK_CAPITAL = &H14
Private Type KeyboardBytes
    kbByte(0 To 255) As Byte
End Type

Private Sub sb_Click(iPanelNumber As Variant)
    If iPanelNumber = 1 Then
        MsgBox "Panel 1 Click"
    End If
End Sub

Private Sub sb_DblClick(iPanelNumber As Variant)
    If iPanelNumber = 2 Then
        sb.PanelCaption(2) = InputBox("Change Caption")
    End If
End Sub

Private Sub sb_MouseDownInPanel(iPanel As Long)
    If iPanel = 5 Then
        MsgBox "Mouse Down In Panel Number " & iPanel
    End If
End Sub

Private Sub Timer1_Timer()
    sb.PanelCaption(1) = Time
End Sub

Public Function GetCapsLockState() As Boolean
'True if caps lock is on
Dim i As Long
    GetKeyboardState kbArray
    i = kbArray.kbByte(VK_CAPITAL)
    If i = 1 Then
        GetCapsLockState = True
    End If
End Function

Private Sub tmrKey_Timer()
    sb.PanelEnabled(6) = GetCapsLockState
End Sub
