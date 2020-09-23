VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Gehe zu"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2415
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zeile:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Call Unload(Me)
End Sub

Private Sub cmdOK_Click()
If IsNumeric(txtLine.Text) Then Call frmMain.Codebox.GotoLine(CLng(txtLine.Text))
Unload Me
End Sub
