VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Mirror Menues!"
   ClientHeight    =   2640
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   5
      Left            =   360
      Picture         =   "Form1.frx":0000
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   4
      Left            =   360
      Picture         =   "Form1.frx":1032
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   3
      Left            =   360
      Picture         =   "Form1.frx":2064
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   2
      Left            =   360
      Picture         =   "Form1.frx":3096
      Top             =   960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":40C8
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgItem 
      Height          =   255
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":50FA
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Menu mirror 
      Caption         =   "Mirror"
      Begin VB.Menu citem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************
'* API Declarations   *
'**********************
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

'**********************
'* Consts             *
'**********************
Const MF_BYCOMMAND = &H0&
Const MF_BITMAP = &H4&

'The procedures of the menu items
Private Sub citem_Click(Index As Integer)
    MsgBox "Click on item #" & CStr(Index + 1)
End Sub

Private Sub Form_Load()
  'Declare Variables
  Dim X%, h1&, h2&, Com&
    'Get the handle of the menu and the handle of its submenu
    h1 = GetMenu(Me.hwnd)
    h2 = GetSubMenu(h1, 0)
    'Create the menu items
    For X = 1 To 5
         Load citem(X)
    Next X
    'Replace the menu entries with bitmaps using the modifymenu function
    For X = 0 To 5
        Com = GetMenuItemID(h2, X)
        Call ModifyMenu(h2, Com, MF_BYCOMMAND Or MF_BITMAP, Com, CLng(imgItem(X).Picture))
    Next X
End Sub
