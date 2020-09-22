VERSION 5.00
Object = "{6C81467A-AA48-11D9-A3BB-0080AD7F0F26}#22.0#0"; "xHTMLLib.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin xHTMLLib.xHTML xHTML1 
      Align           =   1  'Align Top
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   7646
      ScaleMode       =   0
      ScaleWidth      =   11235
      ScaleHeight     =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
    On Error Resume Next
        xHTML1.Height = Me.ScaleHeight
End Sub

