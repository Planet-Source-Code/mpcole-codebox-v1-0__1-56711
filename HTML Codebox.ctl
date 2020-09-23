VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl BSCodebox 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   ScaleHeight     =   5910
   ScaleWidth      =   8340
   Begin VB.PictureBox picTray 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin RichTextLib.RichTextBox rtfCodebox 
         Height          =   3735
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         MaxLength       =   200000
         Appearance      =   0
         RightMargin     =   99999
         TextRTF         =   $"HTML Codebox.ctx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "BSCodebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//BSCodebox (V 1.0)
'//BSCodebox (User Control)
'//Entwicklungsbeginn: 16.08.2004
'//Entwickler: BarbarianSoft - Michael Kull
'//Copyright: 2004 BarbarianSoft
'//////////////////////////////////////////////////////////////////////////////////////

Implements ISubclass

'API Deklarationen
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long

'API Konstanten
Const EM_LINEINDEX = &HBB
Const WM_SETREDRAW = &HB
Const WM_USER = &H400
Const EM_EXGETSEL = (WM_USER + 52)
Const EM_EXSETSEL = (WM_USER + 55)
Const EM_SETSCROLLPOS = (WM_USER + 222)
Const EM_GETSCROLLPOS = (WM_USER + 221)
Const EM_HIDESELECTION = (WM_USER + 63)
Const EM_GETTEXTLENGTHEX = (WM_USER + 95)

Const EM_GETLINECOUNT = &HBA
Const EM_LINELENGTH = &HC1
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_LINEFROMCHAR = &HC9
Const EM_LINESCROLL = &HB6

Const EM_SETTARGETDEVICE = (WM_USER + 72)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type GETTEXTLENGTHEX
    Flags As Long
    Codepage As Integer
End Type

Private Type CHARRANGE
    Min As Long
    Max As Long
End Type

Private ScrollPosition As POINTAPI
Private TextRange As CHARRANGE

'Öffentliche Ereignisse
Event SelChange()
Event Change()
Attribute Change.VB_Description = "Zeigt an, daß sich der Inhalt eines Steuerelements geändert hat."
Event Click()
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Tritt auf, wenn die Maus während einer Drag/Drop-Operation über das Steuerelement bewegt wird, sofern die zugehörige OLEDropMode-Eigenschaft auf ""Manuell"" festgelegt ist."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Tritt auf, wenn eine OLE-Drag/Drop-Operation entweder manuell oder automatisch eingeleitet wird."
Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Tritt im Quellsteuerelement für eine OLE-Drag/Drop-Operation auf, wenn das Ablageziel Daten anfordert, die dem Datenobjekt während des OLEDragStart-Ereignisses nicht zur Verfügung standen."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Tritt im Quellsteuerelement einer OLE-Drag/Drop-Operation auf, wenn der Maus-Cursor geändert werden muß."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event Highlighting()
Event HighlightingCompleted()
Event Modify()

'Standard Eigenschaftswerte
Const DefaultLanguage = "HTML"
Const DefaultSelStart = 0
Const DefaultSelLength = 0
Const DefaultSelText = ""
Const DefaultTextRTF = ""
Const DefaultText = ""
Const DefaultModified = 0
Const DefaultWordWrap = 0
Const DefaultTextLength = 0
Const DefaultFontName = "Courier New"
Const DefaultFontSize = 10

'Eigenschaftsvariablen
Dim ValueLanguage As String
Dim ValueSelStart As Variant
Dim ValueSelLength As Variant
Dim ValueSelText As Variant
Dim ValueText As Variant
Dim ValueModified As Boolean
Dim ValueWordWrap As Boolean
Dim ValueTextLength As Long
Dim ValueFontName As String
Dim ValueFontSize As Integer

'Variablen
Dim Busy As Boolean
Dim LineChanged As Boolean

Private Type Comment
    InComment As Boolean
    CommentStart As Long
    CommentEnd As Long
    CommentClosed As Boolean
End Type
Dim Comment As Comment

Dim VarInScript As Boolean
Dim VarInComment As Boolean
Dim VarInString As Boolean
Dim VarInTag As Boolean

'*** Subclassing ***
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'Response
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
If CurrentMessage = 522 Then
    ISubclass_MsgResponse = emrConsume
Else
    ISubclass_MsgResponse = emrPreprocess
End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim CurrentLine As Long
Dim TP As POINTAPI
Dim PixelsPerLine As Long
Dim ScrollPos As Long

Select Case iMsg
    Case 277
        If wParam = 8 Then
            Call SendMessageLong(rtfCodebox.hwnd, EM_LINESCROLL, 0&, ByVal 0)
        
            ScrollPos = GetScrollPos(rtfCodebox.hwnd, 1)
        
            PixelsPerLine = picTray.TextHeight("IAHÄ90980&^,") / Screen.TwipsPerPixelX
            If Int(ScrollPos / PixelsPerLine) <> (ScrollPos / PixelsPerLine) Then
                Call SendMessage(rtfCodebox.hwnd, EM_GETSCROLLPOS, 0, TP)
                TP.y = TP.y - (ScrollPos Mod PixelsPerLine) + PixelsPerLine
                Call SendMessage(rtfCodebox.hwnd, EM_SETSCROLLPOS, 0&, TP)
            End If
        End If
        
        Call PrintLineNumbers
    Case 522
        PixelsPerLine = picTray.TextHeight("IAHÄ90980&^,") / Screen.TwipsPerPixelX
        
        Call SendMessage(rtfCodebox.hwnd, EM_GETSCROLLPOS, 0, TP)
        
        If wParam > 0 Then
            TP.y = TP.y - (PixelsPerLine * 5)
        Else
            TP.y = TP.y + (PixelsPerLine * 5)
        End If
        
        Call SendMessage(rtfCodebox.hwnd, EM_SETSCROLLPOS, 0&, TP)
        Call PrintLineNumbers
End Select
End Function

Private Sub picTray_GotFocus()
Call rtfCodebox.SetFocus
End Sub

'Die wichtigste Routine für das dynamische Highlighting
Private Sub rtfCodebox_Change()
If Busy = False Then
    RaiseEvent Change
End If
End Sub

Private Sub HighlightDynamic(Optional ParseRanges As Boolean)
'Alte Cursorposition
Dim OldStart As Long
Dim OldLen As Long
'Gewählter Text
Dim StartPos As Long
Dim LineLength As Long
Dim ScriptPos As Long
Dim CommentPos As Long
'Gewählte Zeilen
Dim StartLine As Long

On Error Resume Next
RaiseEvent Highlighting

'Wenn Highlighting bereits aufgefrischt wird, dann abbrechen
If Busy = True Then Exit Sub

Busy = True

If UCase(ValueLanguage) = "TEXT" Then GoTo Enter

'Zunächst die Scroll Position ermitteln
Call SaveScrollPosition

With rtfCodebox
    'Wenn kein Text zum Einfärben vorhanden ist, dann abbrechen
    If .Text = "" Then GoTo Enter
    
    .Locked = True
    
    'Redraw deaktivieren
    'Hab's erst mit WM_SETREDRAW=0 probiert.
    'Dann aber rausgefunden, dass die Verarbeitung mit LOCKWINDOWUPDATE
    'etwa 100% schneller läuft
    'Call SendMessage(.hwnd, WM_SETREDRAW, 0, 0)
    Call LockWindowUpdate(UserControl.hwnd)

    Comment = InComment(.SelStart)
    
    'Zuerst feststellen, ob wir uns in einem Kommentar befinden
    If Comment.InComment Then
        'Wenn wir uns Mitten im Kommentar befinden dann verlassen wir die Routine
        If ParseRanges = False Then
            If Comment.CommentClosed Then GoTo Quit
        End If
        
        'Selektionsdaten für den Kommentar errechnen
        'If UCase(ValueLanguage) <> "HTML" Then GoTo LineSel
        StartPos = Comment.CommentStart
        LineLength = Comment.CommentEnd - Comment.CommentStart
    'Wenn nein, nur die aktuelle Zeile zum Einfärben  auswählen
    Else
        'Selektionsdaten für die gegenwärtige Zeile errechnen
        'StartLine = SendMessage(.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
        StartLine = SendMessageLong(.hwnd, EM_LINEFROMCHAR, .SelStart, 0&)
            
        StartPos = SendMessageLong(.hwnd, EM_LINEINDEX, StartLine, 0&)
        LineLength = SendMessageLong(.hwnd, EM_LINELENGTH, .SelStart, 0&)
        'LineLength = .SelStart - StartPos
        
        'Wir sind zwar nicht unmittelbar in einem Kommentar,
        'wir könnten uns aber in der letzten Zeile eines Kommenateres befinden.
        'In diesem Fall muss der Kommentar unbeding unangetastet bleiben,
        'um das Highlighting nicht zu verlieren. Wir wählen nur den Text
        'nach dem Kommentar End Tag aus!
        If CommentStartChar <> "" Then
            CommentPos = InStrRev(.Text, CommentEndChar, .SelStart, vbTextCompare)
            If CommentPos > StartPos And CommentPos <= StartPos + LineLength Then
                LineLength = LineLength - (CommentPos + (Len(CommentEndChar) - 1) - StartPos)
                StartPos = CommentPos + (Len(CommentEndChar) - 1)
            End If
        End If
    End If

    'Cursorposition speichern
    OldStart = .SelStart
    OldLen = .SelLength

    'Neuen Bereich auswählen
    .SelStart = StartPos
    .SelLength = LineLength
    
    'Farbe auf schwarz setzen
    .SelColor = 0

    'Text einfärben
    If UCase(ValueLanguage) = "HTML" Then
        'Coloring für HTML
        .SelRTF = ColorHTML(.SelText, ValueFontName, ValueFontSize)
    Else
        'Für alle anderen Sprachen
        .SelRTF = ColorCode(.SelText, ValueFontName, ValueFontSize)
    End If

    'Cursorposition wiederherstellen
    .SelStart = OldStart
    .SelLength = OldLen
    
    'Scrollposition setzen
    'Call SendMessage(.hwnd, EM_SETSCROLLPOS, 0&, TempScrollPosition)
    Call RestoreScrollPosition
    Call PrintLineNumbers
    
Quit:
    'Codebox updaten
    'Call SendMessage(.hwnd, WM_SETREDRAW, 1, 0)
    Call LockWindowUpdate(False)
    Call InvalidateRectAsNull(.hwnd, 0&, 0&)
    
    .Locked = False
    .SetFocus
End With

'Neue Scroll Position ermitteln
Call SaveScrollPosition

'Verändert auf Wahr setzen
ValueModified = True

Enter:
Busy = False

RaiseEvent HighlightingCompleted
End Sub

Private Function InComment(SelStart As Long) As Comment
On Error GoTo Error
Dim StartPos As Long
Dim EndPos As Long
Dim SelText As String

With rtfCodebox
    If CommentStartChar = "" Then Exit Function
    If StrComp(Mid(.Text, SelStart - Len(CommentStartChar) + 1, Len(CommentStartChar)), CommentStartChar, vbTextCompare) = 0 Then
        InComment.CommentStart = SelStart - Len(CommentStartChar)
        InComment.CommentEnd = Len(CommentStartChar)
        InComment.CommentClosed = False
        
        If InComment.CommentEnd - InComment.CommentStart <= 0 Then
            InComment.InComment = False
        Else
            InComment.InComment = True
        End If
        Exit Function
    End If
    
    'Kommentarstart ermitteln
    StartPos = InStrRev(.Text, CommentStartChar, SelStart)
    
    If StartPos <= 0 Then Exit Function
    
    If StartPos > 0 Then
        'Text bis zur Cursorposition auswählen
        
        If Mid(.Text, StartPos - Len(CommentChar), Len(CommentChar)) <> CommentChar Then
            SelText = Mid(.Text, StartPos, SelStart - StartPos + 1)
        Else
            SelText = Mid(.Text, StartPos - Len(CommentChar), SelStart - StartPos - Len(CommentChar) + 1)
        End If

        'Sicherstellen, dass keine weiteren Kommentare dazwischenliegen:
        '==> Wir befinden uns DEFINITIV in einem Kommentar!
        If InStrRev(SelText, CommentEndChar, Len(SelText)) <= 0 Then
            If InStr(1, SelText, vbCrLf, vbTextCompare) <= 0 Then
                        
                'Wir markieren den Kommentartext bis zur Cursorposition!
                StartLine = SendMessageLong(.hwnd, EM_LINEFROMCHAR, SelStart, 0&)
            
                InComment.CommentStart = SendMessageLong(.hwnd, EM_LINEINDEX, StartLine, 0&)
                InComment.CommentEnd = SelStart - CommentStart
             
                InComment.CommentClosed = False
                
                If InComment.CommentEnd - InComment.CommentStart <= 0 Then
                    InComment.InComment = False
                Else
                    InComment.InComment = True
                End If
                
            Else
                'Wir markieren den Scripttext bis zur Cursorposition!
                InComment.CommentStart = StartPos - 1
                InComment.CommentEnd = SelStart - StartPos + Len(CommentEndChar) + 1
                InComment.CommentClosed = True
                InComment.InComment = True
            End If
        End If
    End If
End With

Error:
End Function

Private Sub PasteCode(Code As String)
Dim TextRange As String
Dim OldStart As Long

'Abfangen, wenn der Quellcode zu groß ist
If Len(rtfCodebox.Text) + Len(Code) > 200000 Then Exit Sub

LineChanged = True

'Cursorposition speichern
OldStart = rtfCodebox.SelStart

'rtfCodebox.SelColor = 0
rtfCodebox.SelText = Code

If UCase(ValueLanguage) = "HTML" Then
    TextRange = Left(rtfCodebox.Text, 1) & Mid(rtfCodebox.Text, 1, rtfCodebox.SelStart)
    
    If InCommentHTML(TextRange) Or InScriptHTML(TextRange) Or InPropvalHTML(TextRange) Or InTagHTML(TextRange) Then Exit Sub
End If

'Textabschnitt einfärben
Call RefreshRange(OldStart, Len(Code), False)

RaiseEvent Modify
End Sub

Private Sub SaveScrollPosition()
Dim F As GETTEXTLENGTHEX
Const GTL_PRECISE = 2

On Error Resume Next
F.Flags = GTL_PRECISE

Call SendMessage(rtfCodebox.hwnd, EM_EXGETSEL, 0&, TextRange)
Call SendMessage(rtfCodebox.hwnd, EM_GETSCROLLPOS, 0&, ScrollPosition)
End Sub

Private Sub RestoreScrollPosition()
Dim F As GETTEXTLENGTHEX
Const GTL_PRECISE = 2

Call SendMessage(rtfCodebox.hwnd, EM_EXSETSEL, 0&, TextRange)
Call SendMessage(rtfCodebox.hwnd, EM_SETSCROLLPOS, 0&, ScrollPosition)
F.Flags = GTL_PRECISE
End Sub

'###Öffentliche Subroutinen###
Public Sub RefreshHighlighting()
Dim OldStart As Long
Dim OldLen As Long

On Error Resume Next
RaiseEvent Highlighting

Busy = True

'Zunächst die Scroll Position ermitteln
Call SaveScrollPosition

With rtfCodebox
    DoEvents
    
    'Redraw deaktivieren
    Call LockWindowUpdate(UserControl.hwnd)
    'Cursorposition speichern
    OldStart = .SelStart
    OldLen = .SelLength
    
    'Text einfärben
    If UCase(ValueLanguage) = "TEXT" Then
        '.TextRTF = ColorText(rtfCodebox.Text, ValueFontName, ValueFontSize)
        .SelStart = 0
        .SelLength = Len(rtfCodebox.Text)
        .SelColor = 0
        .SelBold = False
    ElseIf UCase(ValueLanguage) = "HTML" Then
        .TextRTF = ColorHTML(rtfCodebox.Text, ValueFontName, ValueFontSize)
    Else
        .TextRTF = ColorCode(rtfCodebox.Text, ValueFontName, ValueFontSize)
    End If

    'Cursorposition wiederherstellen
    .SelStart = OldStart
    .SelLength = OldLen
    
    'Scrollposition setzen
    'Call SendMessage(.hwnd, EM_SETSCROLLPOS, 0&, TempScrollPosition)
    Call RestoreScrollPosition
    Call PrintLineNumbers
    
    'Codebox updaten
    Call LockWindowUpdate(False)
    Call InvalidateRectAsNull(.hwnd, 0&, 0&)
    .SetFocus
End With

'Neue Scroll Position ermitteln
Call SaveScrollPosition

Busy = False

RaiseEvent HighlightingCompleted
End Sub

Private Sub RefreshRange(SelStart As Long, SelLength As Long, SelectRange As Boolean, Optional RestoreSelection As Boolean, Optional OldSelStart As Long, Optional OldSelLength As Long)
Dim OldStart As Long
Dim OldLen As Long

On Error Resume Next
RaiseEvent Highlighting

Busy = True

'Zunächst die Scroll Position ermitteln
Call SaveScrollPosition

With rtfCodebox
    'Redraw deaktivieren
    Call LockWindowUpdate(UserControl.hwnd)
    
    'Cursorposition speichern
    OldStart = SelStart
    OldLen = SelLength
    
    'Neuen Bereich auswählen
    .SelStart = SelStart
    .SelLength = SelLength

    'Farbe auf Schwarz setzen
    .SelColor = 0

    'Text einfärben
    If UCase(ValueLanguage) = "TEXT" Then
        .SelStart = 0
        .SelLength = Len(rtfCodebox.Text)
        .SelColor = 0
    ElseIf UCase(ValueLanguage) = "HTML" Then
        .SelRTF = ColorHTML(.SelText, ValueFontName, ValueFontSize)
    Else
        .SelRTF = ColorCode(.SelText, ValueFontName, ValueFontSize)
    End If

    'Wenn Text eingefügt wird, dann Cursorposition wiederherstellen
    If RestoreSelection = True Then
        .SelStart = OldSelStart
        .SelLength = OldSelLength
    Else
        If SelectRange = True Then
            .SelStart = OldStart
            .SelLength = OldLen
        Else
            .SelStart = .SelStart + .SelLength
            .SelLength = 0
        End If
    End If
    
    'Scrollposition setzen
    Call RestoreScrollPosition
    Call PrintLineNumbers
Quit:
    'Codebox updaten
    Call LockWindowUpdate(False)
    Call InvalidateRectAsNull(.hwnd, 0&, 0&)
    
    .SetFocus
End With

'Neue Scroll Position ermitteln
Call SaveScrollPosition

Busy = False

RaiseEvent HighlightingCompleted
End Sub

Public Sub NewDocument()
rtfCodebox.TextRTF = ""
rtfCodebox.SelColor = 0
End Sub

Public Sub LoadFile(Filename As String)
Dim F As Integer
Dim Result As String

'Existiert die Datei?
If Dir(Filename) <> "" Then
    'Textdatei im Binärmodus öffnen und gesamten Inhalt in einem Rutsch auslesen
    F = FreeFile
    Open Filename For Binary As #F
        Result = Space(LOF(F))
        Get #F, , Result
    Close #F
    
    If Len(Result) > 200000 Then
        Exit Sub
    End If
        
    rtfCodebox.TextRTF = ""
    rtfCodebox.Text = Result
    
    Call PrintLineNumbers
    Call RefreshHighlighting
    
End If
End Sub

Public Sub SaveFile(Filename As String, Format As String)
Open Filename For Output As #1
    'Nur Text speichern
    If UCase(Format) = "TEXT" Then
        Print #1, rtfCodebox.Text
    'Richtext speichern
    ElseIf UCase(Format) = "RTFTEXT" Then
        Print #1, rtfCodebox.TextRTF
    End If
Close #1
End Sub

Public Sub InsertCode(Code As String, Optional SelectRange As Boolean, Optional Refresh As Boolean)
Dim OldStart As Long

'Abfangen, wenn der Quellcode zu groß ist
If Len(rtfCodebox.Text) + Len(Code) > 200000 Then Exit Sub

'Cursorposition speichern
OldStart = rtfCodebox.SelStart

'Text einfärben
If UCase(ValueLanguage) = "TEXT" Then
    rtfCodebox.SelRTF = ColorText(Code, ValueFontName, ValueFontSize)
ElseIf UCase(ValueLanguage) = "HTML" Then
    rtfCodebox.SelRTF = ColorHTML(Code, ValueFontName, ValueFontSize)
Else
    rtfCodebox.SelRTF = ColorCode(Code, ValueFontName, ValueFontSize)
End If

If SelectRange = True Then
    rtfCodebox.SelStart = OldStart
    rtfCodebox.SelLength = Len(Code)
End If

rtfCodebox.SetFocus

RaiseEvent Modify
End Sub

Public Sub Copy()
On Error Resume Next
Call Clipboard.Clear
Call Clipboard.SetText(rtfCodebox.SelText)
End Sub

Public Sub Cut()
On Error Resume Next
Call Clipboard.Clear
Call Clipboard.SetText(rtfCodebox.SelText)
rtfCodebox.SelText = ""
End Sub

Public Sub Paste()
On Error Resume Next
Call PasteCode(Clipboard.GetText)
End Sub

Public Sub SelectAll()
On Error Resume Next
rtfCodebox.SelStart = 0
rtfCodebox.SelLength = Len(rtfCodebox.Text)
End Sub

Public Sub Deselect()
On Error Resume Next
rtfCodebox.SelLength = 0
End Sub

Public Sub SelectAbove()
Dim SelStart As Long

On Error Resume Next
SelStart = rtfCodebox.SelStart
rtfCodebox.SelStart = 0
rtfCodebox.SelLength = SelStart
End Sub

Public Sub SelectBelow()
On Error Resume Next
rtfCodebox.SelStart = rtfCodebox.SelStart
rtfCodebox.SelLength = Len(rtfCodebox.Text) - rtfCodebox.SelStart
End Sub

Public Sub SelectLine()
Dim StartLine As Long
Dim StartPos As Long
Dim LineLength As Long

On Error Resume Next
StartLine = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart, 0&)
            
StartPos = SendMessageLong(rtfCodebox.hwnd, EM_LINEINDEX, StartLine, 0&)
LineLength = SendMessageLong(rtfCodebox.hwnd, EM_LINELENGTH, rtfCodebox.SelStart, 0&)

rtfCodebox.SelStart = StartPos
rtfCodebox.SelLength = LineLength
End Sub

Public Sub GotoLine(Line As Long)
On Error Resume Next
StartPos = SendMessageLong(rtfCodebox.hwnd, EM_LINEINDEX, Line - 1, 0&)

rtfCodebox.SelStart = StartPos
rtfCodebox.SelLength = 0
End Sub

Public Sub Delete()
On Error Resume Next
rtfCodebox.SelText = ""

RaiseEvent Modify
End Sub

Public Sub SetCursor(Position As Long)
On Error Resume Next
rtfCodebox.SelStart = Position
rtfCodebox.SelLength = 0
rtfCodebox.SetFocus
End Sub

Public Function Find(FindString As String, Optional Start As Long, Optional vEnd As Long, Optional Options As Long) As Long
Find = rtfCodebox.Find(FindString, Start, vEnd, Options)
End Function

'###Öffentliche Ereignisse###
Private Sub rtfCodebox_Click()
RaiseEvent Click
End Sub

Private Sub rtfCodebox_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error

If UCase(ValueLanguage) <> "HTML" Then
    Select Case KeyCode
        Case 38, 40
            If LineChanged = True Then
                Call HighlightDynamic(True)
                LineChanged = False
            End If
        Case 13
            If LineChanged = True Then
                Call HighlightDynamic(True)
                LineChanged = False
            End If
        Case Else
            LineChanged = True
    End Select
Else

End If
Error:

If Shift = vbCtrlMask Then
    Select Case KeyCode
        Case vbKeyX
            Call Cut
            KeyCode = 0
        Case vbKeyC
            Call Copy
            KeyCode = 0
        Case vbKeyV
            Call Paste
            KeyCode = 0
        Case vbKeyZ
            KeyCode = 0
    End Select
Else
    Select Case KeyCode
        Case vbKeyF5
            Call RefreshHighlighting
        Case vbKeyTab
            rtfCodebox.SelText = vbTab & rtfCodebox.SelText
            KeyCode = 0
        Case 46, 8
            RaiseEvent Modify
    End Select
End If

RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub rtfCodebox_KeyPress(KeyAscii As Integer)
Dim Char As String
Dim InTag As Boolean

If UCase(ValueLanguage) = "HTML" Then
    Call LockWindowUpdate(UserControl.hwnd)
    
    If VarInScript Or VarInComment Then
        Call LockWindowUpdate(False)
        rtfCodebox.Locked = False
        Exit Sub
    Else
        If VarInTag = False Then rtfCodebox.SelColor = TextColor
        
        If KeyAscii = 60 Then
            rtfCodebox.SelColor = KeywordColor(1)
        ElseIf KeyAscii = 62 Then
            rtfCodebox.SelColor = TextColor
        End If
    End If

    If VarInTag Then
        If Chr(KeyAscii) = "-" Then
            If Not Len(rtfCodebox.Text) < 3 Then
                If Mid(rtfCodebox.Text, rtfCodebox.SelStart - 2, 3) = "<!-" Then
                    rtfCodebox.Locked = True
                    rtfCodebox.SelStart = rtfCodebox.SelStart - 3
                    rtfCodebox.SelLength = 3
                    rtfCodebox.SelColor = CommentColor
                    rtfCodebox.SelStart = rtfCodebox.SelStart + 4
                    rtfCodebox.Locked = False
                End If
            End If
        End If
        
        If Chr(KeyAscii) = "t" Then
            If Not Len(rtfCodebox.Text) < 6 Then
                If Mid(rtfCodebox.Text, rtfCodebox.SelStart - 5, 6) = "<scrip" Then
                    rtfCodebox.Locked = True
                    rtfCodebox.SelStart = rtfCodebox.SelStart - 6
                    rtfCodebox.SelLength = 6
                    rtfCodebox.SelColor = ScriptColor
                    rtfCodebox.SelStart = rtfCodebox.SelStart + 6
                    rtfCodebox.Locked = False
                End If
            End If
        End If
        
        If KeyAscii = 32 Then
            If VarInString Then
                rtfCodebox.SelColor = QuotationColor
            Else
                rtfCodebox.SelColor = KeywordColor(1)
            End If
        ElseIf Chr(KeyAscii) = "=" Then
            rtfCodebox.SelText = "="
            rtfCodebox.SelColor = QuotationColor
            KeyAscii = 0
        ElseIf Chr(KeyAscii) = ">" Then
            rtfCodebox.SelColor = KeywordColor(1)
            rtfCodebox.SelText = ">"
            KeyAscii = 0
            rtfCodebox.SelColor = TextColor
        End If
    End If
    
    
    Call LockWindowUpdate(False)
End If

Error:
RaiseEvent KeyPress(KeyAscii)
RaiseEvent Modify
End Sub

Private Function InTagHTML(Range As String) As Boolean
On Error GoTo Error
DoEvents
If rtfCodebox.SelStart > 0 Then
    If InStrRev(Range, "<") > InStrRev(Range, ">") Then InTagHTML = True
End If

Error:
End Function

Private Function InCommentHTML(Range As String) As Boolean
On Error GoTo Error
DoEvents
If rtfCodebox.SelStart > 0 Then
    If InStrRev(Range, "<!--") > InStrRev(Range, "-->") Then InCommentHTML = True
End If

Error:
End Function

Private Function InScriptHTML(Range As String) As Boolean
On Error GoTo Error
DoEvents
If rtfCodebox.SelStart > 0 Then
    If InStrRev(Range, "<script") > InStrRev(Range, "</script>") Then InScriptHTML = True
End If

Error:
End Function

Private Function InPropvalHTML(Range As String) As Boolean
Dim x As Long
Dim y As Long

On Error GoTo Error
DoEvents
  
InPropvalHTML = False
With rtfCodebox
    x = InStrRev(Range, Chr(34))
    y = InStrRev(Range, "=")
    
    If x > y Then
        If TrimWS(Mid(Range, y, x - y)) = "=" Then
            InPropvalHTML = True
        End If
    ElseIf y = Len(Range) + 1 Then
        InPropvalHTML = True
    End If
End With

Error:
End Function

Private Sub rtfCodebox_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rtfCodebox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If UCase(ValueLanguage) <> "HTML" Then
    If LineChanged = True Then
        Call HighlightDynamic
        LineChanged = False
    End If
End If

RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtfCodebox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtfCodebox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub rtfCodebox_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub rtfCodebox_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub rtfCodebox_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub rtfCodebox_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub rtfCodebox_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub rtfCodebox_SelChange()
Dim TextRange As String

Call PrintLineNumbers

If UCase(ValueLanguage) = "HTML" Then
    TextRange = Left(rtfCodebox.Text, 1) & Mid(rtfCodebox.Text, 1, rtfCodebox.SelStart)
    
    VarInComment = InCommentHTML(TextRange)
    VarInScript = InScriptHTML(TextRange)
    VarInString = InPropvalHTML(TextRange)
    VarInTag = InTagHTML(TextRange)
End If

If Busy = False Then
    RaiseEvent SelChange
End If
End Sub

'###Initialisierung###
Private Sub UserControl_EnterFocus()
Call rtfCodebox.SetFocus

Call PrintLineNumbers
End Sub

Private Sub UserControl_Initialize()
Call AttachMessage(Me, rtfCodebox.hwnd, 277)
Call AttachMessage(Me, rtfCodebox.hwnd, 522)
Call AttachMessage(Me, rtfCodebox.hwnd, 257)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

Call picTray.Move(0, 0, UserControl.Width, UserControl.Height)
Call rtfCodebox.Move(720, 0, picTray.Width - 780, picTray.Height - 60)

Call PrintLineNumbers
End Sub

'###Öffentliche Eigenschaften###
'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
ValueLanguage = DefaultLanguage
ValueSelStart = DefaultSelStart
ValueSelLength = DefaultSelLength
ValueSelText = DefaultSelText
ValueText = DefaultText
ValueModified = DefaultModified
ValueTextLength = DefaultTextLength
ValueFontName = DefaultFontName
ValueFontSize = DefaultFontSize
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
ValueLanguage = PropBag.ReadProperty("Language", DefaultLanguage)
ValueSelStart = PropBag.ReadProperty("SelStart", DefaultSelStart)
ValueSelLength = PropBag.ReadProperty("SelLength", DefaultSelLength)
ValueSelText = PropBag.ReadProperty("SelText", DefaultSelText)
ValueText = PropBag.ReadProperty("Text", DefaultText)
ValueModified = PropBag.ReadProperty("Modified", DefaultModified)
rtfCodebox.Locked = PropBag.ReadProperty("Locked", Falsch)
rtfCodebox.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", Falsch)
rtfCodebox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
picTray.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
rtfCodebox.Text = PropBag.ReadProperty("TextRTF", DefaultTextRTF)
rtfCodebox.HideSelection = PropBag.ReadProperty("HideSelection", Falsch)
ValueTextLength = PropBag.ReadProperty("TextLength", DefaultTextLength)
ValueFontName = PropBag.ReadProperty("FontName", DefaultFontName)
ValueFontSize = PropBag.ReadProperty("FontSize", DefaultFontSize)
End Sub

Private Sub UserControl_Terminate()
Call DetachMessage(Me, rtfCodebox.hwnd, 277)
Call DetachMessage(Me, rtfCodebox.hwnd, 522)
Call DetachMessage(Me, rtfCodebox.hwnd, 257)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Language", ValueLanguage, DefaultLanguage)
Call PropBag.WriteProperty("SelStart", ValueSelStart, DefaultSelStart)
Call PropBag.WriteProperty("SelLength", ValueSelLength, DefaultSelLength)
Call PropBag.WriteProperty("SelText", ValueSelText, DefaultSelText)
Call PropBag.WriteProperty("Modified", ValueModified, DefaultModified)
Call PropBag.WriteProperty("Text", ValueText, DefaultText)
Call PropBag.WriteProperty("Locked", rtfCodebox.Locked, Falsch)
Call PropBag.WriteProperty("AutoVerbMenu", rtfCodebox.AutoVerbMenu, Falsch)
Call PropBag.WriteProperty("BackColor", rtfCodebox.BackColor, &H80000005)
Call PropBag.WriteProperty("BorderStyle", picTray.BorderStyle, 1)
Call PropBag.WriteProperty("TextRTF", rtfCodebox.Text, DefaultTextRTF)
Call PropBag.WriteProperty("HideSelection", rtfCodebox.HideSelection, Falsch)
Call PropBag.WriteProperty("TextLength", ValueTextLength, DefaultTextLength)
Call PropBag.WriteProperty("FontName", ValueFontName, DefaultFontName)
Call PropBag.WriteProperty("FontSize", ValueFontSize, DefaultFontSize)
End Sub

Public Property Get Language() As String
Language = ValueLanguage
End Property

Public Property Let Language(ByVal NewLanguage As String)
Dim OldStart As Long
Dim OldLen As Long

ValueLanguage = NewLanguage

'Hier wird das Syntax Schema geladen
Call LoadSyntax(ValueLanguage)

Call RefreshHighlighting

PropertyChanged "Language"
End Property

Public Property Get LineLength() As Long
Attribute LineLength.VB_MemberFlags = "400"
LineLength = SendMessageLong(rtfCodebox.hwnd, EM_LINELENGTH, rtfCodebox.SelStart, 0&)
End Property

Public Property Let LineLength(ByVal NewLineLength As Long)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get LineCount() As Long
Attribute LineCount.VB_MemberFlags = "400"
LineCount = SendMessageLong(rtfCodebox.hwnd, EM_GETLINECOUNT, 0&, 0&)
End Property

Public Property Let LineCount(ByVal NewLineCount As Long)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get CurrentLine() As Long
Attribute CurrentLine.VB_MemberFlags = "400"
CurrentLine = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart, 0&)
End Property

Public Property Get CurrentCol() As Long
Dim StartLine As Long

StartLine = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart, 0&)
CurrentCol = rtfCodebox.SelStart - SendMessageLong(rtfCodebox.hwnd, EM_LINEINDEX, StartLine, 0&)
End Property

Public Property Let CurrentCol(ByVal NewCurrentCol As Long)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get FirstVisibleLine() As Long
Attribute FirstVisibleLine.VB_MemberFlags = "400"
FirstVisibleLine = SendMessageLong(rtfCodebox.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Property

Public Property Let FirstVisibleLine(ByVal NewFirstVisibleLine As Long)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get SelStart() As Long
SelStart = rtfCodebox.SelStart
End Property

Public Property Let SelStart(ByVal NewSelStart As Long)
rtfCodebox.SelStart = NewSelStart
PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Long
SelLength = rtfCodebox.SelLength
End Property

Public Property Let SelLength(ByVal NewSelLength As Long)
rtfCodebox.SelLength = NewSelLength

PropertyChanged "SelLength"
End Property

Public Property Get SelText() As String
SelText = rtfCodebox.SelText
End Property

Public Property Let SelText(ByVal NewSelText As String)
rtfCodebox.SelText = NewSelText
If Len(rtfCodebox.Text) + Len(NewSelText) > 200000 Then Exit Property
PropertyChanged "SelText"
End Property

Public Property Get Text() As String
Text = rtfCodebox.Text
End Property

Public Property Let Text(ByVal NewText As String)
If Len(NewText) > 200000 Then Exit Property
Busy = True
rtfCodebox.Text = NewText
Call RefreshHighlighting
Busy = False
PropertyChanged "Text"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Legt einen Wert fest, der anzeigt, ob der Inhalt eines RTF-Steuerelements bearbeitet werden kann, oder gibt diesen Wert zurück."
Locked = rtfCodebox.Locked
End Property

Public Property Let Locked(ByVal NewLocked As Boolean)
rtfCodebox.Locked() = NewLocked
PropertyChanged "Locked"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Legt die Hintergrundfarbe eines Objekts fest oder gibt diese zurück."
BackColor = rtfCodebox.BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
rtfCodebox.BackColor() = NewBackColor
PropertyChanged "BackColor"
End Property

Public Property Get Modified() As Boolean
Modified = ValueModified
End Property

Public Property Let Modified(ByVal NewModified As Boolean)
ValueModified = NewModified
PropertyChanged "Modified"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
BorderStyle = picTray.BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As Integer)
picTray.BorderStyle() = NewBorderStyle
PropertyChanged "BorderStyle"
End Property

Public Property Get TextRTF() As String
TextRTF = rtfCodebox.TextRTF
End Property

Public Property Let TextRTF(ByVal NewTextRTF As String)
rtfCodebox.TextRTF() = NewTextRTF
PropertyChanged "TextRTF"
End Property

Public Property Get SelRTF() As String
SelRTF = rtfCodebox.SelRTF
End Property

Public Property Let SelRTF(ByVal NewSelRTF As String)
rtfCodebox.SelRTF() = NewSelRTF
PropertyChanged "SelRTF"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Legt eine Wert fest, der angibt, ob das ausgewählte Element hervorgehoben bleibt, wenn das Steuerelement den Fokus abgibt, oder gibt diesen Wert zurück."
HideSelection = rtfCodebox.HideSelection
End Property

Public Property Let HideSelection(ByVal NewHideSelection As Boolean)
rtfCodebox.HideSelection() = NewHideSelection
PropertyChanged "HideSelection"
End Property

Public Property Get WordWrap() As Boolean
WordWrap = ValueWordWrap
End Property

Public Property Let WordWrap(ByVal NewWordWrap As Boolean)
ValueWordWrap = NewWordWrap

If ValueWordWrap = True Then
    Call SendMessageLong(rtfCodebox.hwnd, EM_SETTARGETDEVICE, 0&, 0&)
Else
    Call SendMessageLong(rtfCodebox.hwnd, EM_SETTARGETDEVICE, 0&, 1)
End If

PropertyChanged "WordWrap"
End Property

Public Property Get TextLength() As Long
TextLength = Len(rtfCodebox.Text)
End Property

Public Property Let TextLength(ByVal NewTextLength As Long)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Legt einen Wert fest oder gibt einen Wert zurück, der anzeigt, ob es eine maximale Anzahl an Zeichen gibt, die ein RTF-Steuerelement aufnehmen kann, und gibt gegebenenfalls diese Zahl an."
hwnd = rtfCodebox.hwnd
End Property

Public Property Get CanCopy() As Boolean
Attribute CanCopy.VB_MemberFlags = "400"
On Error Resume Next

If Busy = True Then Exit Property
If rtfCodebox.SelText <> "" Then
    CanCopy = True
Else
    CanCopy = False
End If
End Property

Public Property Let CanCopy(ByVal New_CanCopy As Boolean)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get CanPaste() As Boolean
Attribute CanPaste.VB_MemberFlags = "400"
On Error Resume Next

If Busy = True Then Exit Property
If Clipboard.GetText <> "" Then
    CanPaste = True
Else
    CanPaste = False
End If
End Property

Public Property Let CanPaste(ByVal New_CanPaste As Boolean)
If Ambient.UserMode = False Then Err.Raise 387
If Ambient.UserMode Then Err.Raise 382
End Property

Public Property Get FontName() As String
FontName = ValueFontName
End Property

Public Property Let FontName(ByVal NewFontName As String)
ValueFontName = NewFontName
PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Integer
FontSize = ValueFontSize
End Property

Public Property Let FontSize(ByVal NewFontSize As Integer)
ValueFontSize = NewFontSize
PropertyChanged "FontSize"
End Property

Private Sub PrintLineNumbers()
Dim i As Integer
Dim StartLine As Integer
Dim EndLine As Integer
Dim LineCount As Integer
Dim StartSel As Integer
Dim EndSel As Integer

picTray.Cls

StartLine = SendMessageLong(rtfCodebox.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&) + 1
LineCount = SendMessageLong(rtfCodebox.hwnd, EM_GETLINECOUNT, 0&, 0&)

picTray.FontName = rtfCodebox.Font.Name
picTray.FontSize = rtfCodebox.Font.Size

EndLine = (rtfCodebox.Height / picTray.TextHeight("Xyzi^")) + StartLine

If EndLine > LineCount Then EndLine = LineCount

For i = StartLine To EndLine
    
    StartSel = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart, 0&) + 1
    If rtfCodebox.SelLength > 0 Then
        EndSel = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart + rtfCodebox.SelLength - 1, 0&) + 1
    Else
        EndSel = SendMessageLong(rtfCodebox.hwnd, EM_LINEFROMCHAR, rtfCodebox.SelStart + rtfCodebox.SelLength, 0&) + 1
    End If
    
    If i >= StartSel And i <= EndSel Then
        picTray.FontBold = True
        picTray.ForeColor = vbBlack
    Else
        picTray.FontBold = False
        picTray.ForeColor = &H808080
    End If
    
    picTray.Print Chr(32) & Format(i, "0000")
Next i
End Sub

