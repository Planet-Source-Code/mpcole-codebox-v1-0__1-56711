VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Test"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8310
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin Test.BSCodebox Codebox 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   300
      Left            =   7320
      TabIndex        =   0
      Top             =   5760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   6960
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar Statusbar 
      Align           =   2  'Unten ausrichten
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6540
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   556
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "FEST"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "EINFG"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "ROLL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Zeile: 1"
            TextSave        =   "Zeile: 1"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Spalte: 0"
            TextSave        =   "Spalte: 0"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Zeilenlänge: 0"
            TextSave        =   "Zeilenlänge: 0"
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuFile 
      Caption         =   "Datei"
      Begin VB.Menu mnuNew 
         Caption         =   "Neu"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Öffnen"
         Shortcut        =   ^O
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Speichern"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Speichern unter"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Beenden"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Bearbeiten"
      Begin VB.Menu mnuUndo 
         Caption         =   "Rückgängig"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Wiederholen"
         Shortcut        =   ^Y
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Ausschneiden"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Kopieren"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Einfügen"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Löschen"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectUp 
         Caption         =   "Oberhalb markieren"
      End
      Begin VB.Menu mnuSelectDown 
         Caption         =   "Unterhalb markieren"
      End
      Begin VB.Menu mnuSelectLine 
         Caption         =   "Zeile markieren"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Alles markieren"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDeselect 
         Caption         =   "Markierung aufheben"
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "Gehe zu"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Suchen und Ersetzen"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnuFont 
         Caption         =   "Schriftart"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "Ansicht"
      Begin VB.Menu mnuSyntax 
         Caption         =   "Text"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hilfe"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFile As String

Private Sub Codebox_Modify()
On Error Resume Next
If Mid(Me.Caption, Len(Me.Caption) - 1, 1) <> "*" Then
Me.Caption = "Test - [" & GetFileName(CurrentFile) & "*]"
End If
End Sub

Private Sub Codebox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call Me.PopupMenu(mnuEdit)
End If
End Sub

Private Sub Codebox_SelChange()
On Error Resume Next
Statusbar.Panels(6).Text = "Zeile: " & Codebox.CurrentLine + 1
Statusbar.Panels(7).Text = "Spalte: " & Codebox.CurrentCol
Statusbar.Panels(8).Text = "Zeilenlänge: " & Codebox.LineLength
End Sub

Private Sub Form_Load()
Call LoadLanguages

Call mnuNew_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
Call Codebox.Move(0, 0, Me.Width - 100, Me.Height - 1000)
Call prgBar.Move(30, Me.Height - 950, 1410, 240)
End Sub

Private Sub mnuCopy_Click()
Call Codebox.Copy
End Sub

Private Sub mnuCut_Click()
Call Codebox.Cut
End Sub

Private Sub mnuDelete_Click()
Call Codebox.Delete
End Sub

Private Sub mnuDeselect_Click()
Call Codebox.Deselect
End Sub

Private Sub mnuFind_Click()
Call frmFind.Show(, Me)
End Sub

Private Sub mnuGoto_Click()
Call frmGoto.Show(vbModal)
End Sub

Private Sub mnuNew_Click()
CurrentFile = "Unbenannt"

Call Codebox.NewDocument

Me.Caption = "Test - [Unbenannt]"
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Error
Cmd.Filter = "Alle unterstützten Dateien|*.htm;*.html;*.aps;*.css;*.js;*.xml;*.vbs;*.php;*.txt|HTML Datei (*.htm;*.html)|*.htm;*.html|PHP Datei (*.php)|*.php|ASP Datei (*.asp)|*.asp|CSS Datei (*.css)|*.css|Java Script (*.js)|*.js|VB Script (*.vbs)|*.vbs|XML Datei (*.xml)|*.xml|Textdatei (*.txt)|*.txt"
Cmd.FilterIndex = 1
Cmd.ShowOpen

CurrentFile = Cmd.Filename

Me.Caption = "Test - [" & Cmd.FileTitle & "]"

Call Codebox.LoadFile(Cmd.Filename)

Undo.Reset
Error:
End Sub

Private Sub mnuPaste_Click()
Call Codebox.Paste
End Sub

Private Sub mnuQuit_Click()
Call Unload(Me)
End Sub

Private Sub mnuSave_Click()
If CurrentFile = "Unbenannt" Then
    mnuSaveAs_Click
Else
    If Mid(CurrentFile, 3, 1) = "\" Then Call Codebox.SaveFile(CurrentFile, "Text")
    Me.Caption = "Test - [" & Cmd.FileTitle & "]"
End If
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo Error
Cmd.Filter = "HTML Datei (*.htm;*.html)|*.htm;*.html|PHP Datei (*.php)|*.php|ASP Datei (*.asp)|*.asp|CSS Datei (*.css)|*.css|Java Script (*.js)|*.js|VB Script (*.vbs)|*.vbs|XML Datei (*.xml)|*.xml|Textdatei (*.txt)|*.txt"
Cmd.ShowSave

Call Codebox.SaveFile(Cmd.Filename, "Text")

Me.Caption = "Test - [" & Cmd.FileTitle & "]"

Error:
End Sub

Private Sub mnuSelectAll_Click()
Call Codebox.SelectAll
End Sub

Private Sub mnuSelectDown_Click()
Call Codebox.SelectBelow
End Sub

Private Sub mnuSelectLine_Click()
Call Codebox.SelectLine
End Sub

Private Sub mnuSelectUp_Click()
Call Codebox.SelectAbove
End Sub

Private Sub LoadLanguages()
Dim Filename As String
Dim Checked As Boolean
Dim i As Integer

Filename = Dir(App.Path & "\syntax\*.stx")
If Filename <> "" Then
    Call Load(mnuSyntax(1))
    mnuSyntax(1).Caption = Left(Filename, Len(Filename) - 4)
        
    If mnuSyntax(i).Caption = "HTML" Then
        mnuSyntax(1).Checked = True
        Call mnuSyntax_Click(i)
        Checked = True
    End If
End If

For i = 2 To 13
    Filename = Dir
    
    If Filename = "" Then Exit For
    
    Call Load(mnuSyntax(i))
    mnuSyntax(i).Caption = Left(Filename, Len(Filename) - 4)
    
    If mnuSyntax(i).Caption = "HTML" Then
        mnuSyntax(i).Checked = True
        Call mnuSyntax_Click(i)
        Checked = True
    End If
Next i

If Checked = False Then mnuSyntax(0).Checked = True
End Sub

Private Sub mnuSyntax_Click(Index As Integer)
Dim i As Integer

For i = 0 To mnuSyntax.UBound
    mnuSyntax(i).Checked = False
Next i

mnuSyntax(Index).Checked = True

Codebox.Language = mnuSyntax(Index).Caption
End Sub

Public Function GetFileName(Filename As String) As String
GetFileName = Mid(Filename, InStrRev(Filename, "\") + 1)
End Function
