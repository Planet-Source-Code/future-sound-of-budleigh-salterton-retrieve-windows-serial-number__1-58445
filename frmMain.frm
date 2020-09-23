VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Serial Number"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":61CA
   ScaleHeight     =   1365
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSerial 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   330
      Left            =   2025
      TabIndex        =   0
      Top             =   45
      Width           =   5100
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    If App.PrevInstance Then End
    lblSerial.Caption = QueryValue(HKEY_LOCAL_MACHINE, "software\microsoft\windows\currentversion", "Productkey")
    
End Sub
