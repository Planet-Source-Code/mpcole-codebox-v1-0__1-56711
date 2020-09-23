VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Suchen und Ersetzen"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkCase 
      Caption         =   "Groß- und Kleinschreibung beachten"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Weiter&suchen"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "A&bbrechen"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Ersetzen"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "&Alle ersetzen"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Nur ganzes Wort suchen"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox txtFind 
      Height          =   315
      ItemData        =   "frmFind.frx":000C
      Left            =   960
      List            =   "frmFind.frx":01BA
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.ComboBox txtReplace 
      Height          =   315
      ItemData        =   "frmFind.frx":06A9
      Left            =   960
      List            =   "frmFind.frx":0857
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblLabel1 
      Caption         =   "Suchen:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblLabel2 
      Caption         =   "Ersetzen:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastPosition As Long

Private Sub cmdCancel_Click()
Call Unload(Me)
End Sub

Private Sub cmdFind_Click()
Dim Options As Long

If chkWholeWord.Value = 1 Then Options = Options + rtfWholeWord
If chkCase.Value = 1 Then Options = Options + rtfMatchCase

LastPosition = frmMain.Codebox.Find(txtFind.Text, LastPosition + 1, Len(frmMain.Codebox.Text), Options)

If LastPosition = -1 Then
    Call MsgBox("Der angegebene Bereich wurde durchsucht.", vbExclamation, "Para Essentials")
    LastPosition = -1
End If
End Sub

Private Sub cmdReplace_Click()
Dim Options As Long

If chkWholeWord.Value = 1 Then Options = Options + rtfWholeWord
If chkCase.Value = 1 Then Options = Options + rtfMatchCase

If LastPosition <> -1 And frmMain.Codebox.SelLength >= 1 Then frmMain.Codebox.SelText = txtReplace.Text

LastPosition = frmMain.Codebox.Find(txtFind.Text, LastPosition + 1, Len(frmMain.Codebox.Text), Options)

If LastPosition = -1 Then
    Call MsgBox("Der angegebene Bereich wurde durchsucht.", vbExclamation, "Para Essentials")
    LastPosition = -1
End If
End Sub

Private Sub cmdReplaceAll_Click()
Dim Options As Long

If chkWholeWord.Value = 1 Then Options = Options + rtfWholeWord
If chkCase.Value = 1 Then Options = Options + rtfMatchCase

Do
    LastPosition = frmMain.Codebox.Find(txtFind.Text, LastPosition + 1, Len(frmMain.Codebox.Text), Options)
    
    If LastPosition = -1 Then
        Call MsgBox("Der angegebene Bereich wurde durchsucht.", vbExclamation, "Para Essentials")
        LastPosition = -1
        Exit Do
    Else
        frmMain.Codebox.SelText = txtReplace.Text
    End If
Loop Until LastPosition = -1
End Sub

