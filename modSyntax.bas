Attribute VB_Name = "modSyntax"
'//////////////////////////////////////////////////////////////////////////////////////
'//BSCodebox (V 1.0)
'//modSyntax (Syntax Highlighting und Syntax Scheme Routinen)
'//Entwicklungsbeginn: 16.08.2004
'//Entwickler: BarbarianSoft - Michael Kull
'//Copyright: 2004 BarbarianSoft
'//////////////////////////////////////////////////////////////////////////////////////

'API Deklarationen
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Öffentliche Options Variablen
Public ProperCase As Byte
Public CommentChar As String
Public CommentStartChar As String
Public CommentEndChar As String
Public KeywordsBold As Byte

'Farb Variablen
Public KeywordColor(1 To 2) As Long
Public OperatorColor As Long
Public CommentColor As Long
Public QuotationColor As Long
Public TextColor As Long
Public ScriptColor As Long

'Keyword Variablen
Dim Keywords As String
Dim KeywordsAlt As String
Dim Operator As String

'Variablen
Dim Delimiter As String
Dim Delimiters(30) As String

Public Function LoadSyntax(Language As String)
Dim i As Integer
On Error Resume Next
DoEvents

'Werte von Syntax Schema einlesen
Extensions = TrimWS(LoadIni(App.Path & "\syntax\" & Language & ".stx", "General", "Extensions"))

ProperCase = CInt(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Options", "ProperCase"))
CommentChar = TrimWS(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Options", "CommentChar"))
CommentStartChar = TrimWS(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Options", "CommentStart"))
CommentEndChar = TrimWS(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Options", "CommentEnd"))
KeywordsBold = CInt(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Options", "KeywordsBold"))

Keywords = Chr(1) & Chr(32) & LoadIni(App.Path & "\syntax\" & Language & ".stx", "Syntax", "Keywords") & Chr(1)
KeywordsAlt = Chr(1) & Chr(32) & LoadIni(App.Path & "\syntax\" & Language & ".stx", "Syntax", "KeywordsAlt") & Chr(1)
Operator = LoadIni(App.Path & "\syntax\" & Language & ".stx", "Syntax", "Operators")

CommentColor = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "Comments"))
QuotationColor = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "Strings"))
ScriptColor = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "Scripts"))
KeywordColor(1) = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "Keywords"))
KeywordColor(2) = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "KeywordsAlt"))
OperatorColor = CLng(LoadIni(App.Path & "\syntax\" & Language & ".stx", "Colors", "Operators"))

'String Delimiter einlesen
Delimiter = Chr(34) & vbTab & ",(){}[]+*%/='~!&|<>?:;.#@´`¨¯"

For i = 0 To Len(Delimiter) - 1
    Delimiters(i) = Mid(Delimiter, i + 1, 1)
    Select Case Delimiters(i)
        Case "\"
            Delimiters(i) = "\\"
        Case "}"
            Delimiters(i) = "\}"
        Case "{"
            Delimiters(i) = "\{"
    End Select
Next i
End Function

Public Function ColorCode(Text As String, FontName As String, FontSize As Integer) As String
Dim Lines() As String
Dim Line As String
Dim Words() As String
Dim Pos As Integer
Dim Word As String

Dim Operators() As String

Dim RTF As String
Dim ALLRTF As String
Dim Header As String

Dim onSingleComment As Boolean
Dim onComment As Boolean
Dim onQuotation As Boolean
Dim onTag As Boolean

Dim i As Integer
Dim j As Integer
Dim x As Integer

Dim onBold As String
Dim offBold As String

'Bold Values setzen
If KeywordsBold = 1 Then
    onBold = "\b1" & Chr(32)
    offBold = "\b0"
Else
    onBold = Chr(32)
    offBold = ""
End If

'Liste der Operatoren erstellen
Operators = Split(Operator, Chr(32))

'Text in Zeilen splitten
Lines = Split(Text, vbCrLf)

For i = 0 To UBound(Lines)
    DoEvents
    If Len(Text) > 40000 Then frmMain.prgBar.Value = i / (UBound(Lines) / 100)
    
    'Variablen zurücksetzen
    onSingleComment = False
    onQuotation = False
    
    'Einzelne Linie aus Array
    Line = Lines(i)
    
    'RTF Zeichen maskieren
    Line = Replace(Line, "\", "\\")
    Line = Replace(Line, "}", "\}")
    Line = Replace(Line, "{", "\{")
    
    'Ersetzte Kommentarzeichen
    Line = Replace(Line, CommentChar, "¨", , , vbTextCompare)
    Line = Replace(Line, CommentStartChar, "´", , , vbTextCompare)
    Line = Replace(Line, CommentEndChar, "`", , , vbTextCompare)
    
    'Delimiter mit Leerzeichen trennen
    For j = 0 To UBound(Delimiters)
        Line = Replace(Line, Delimiters(j), Delimiters(j) & Chr(32))
    Next j

    'Zeile in Wörter splitten
    Words = Split(Line)
    
    'Für alle Wörter in Zeile
    For j = 0 To UBound(Words)
        Select Case UCase(Words(j))
            'Einfacher Kommentar
            Case "¨"
                If onQuotation = False Then
                    If onSingleComment = False Then
                        onSingleComment = True
                        Words(j) = "\cf4" & Chr(32) & Words(j)
                        
                        GoTo EndLine
                    End If
                End If
            'Kommentar Anfang
            Case "´"
                If onQuotation = False Then
                    If onComment = False Then
                        onComment = True
                        Words(j) = "\cf4" & Chr(32) & Words(j)
                    End If
                End If
            'Kommentar Ende
            Case "`"
                If onQuotation = False Then
                    If onComment = True Then
                        onComment = False
                        Words(j) = Words(j) & "\cf0"
                    End If
                End If
            'Strings
            Case Chr(34)
                If onComment = False And onSingleComment = False Then
                    If onQuotation = False Then
                        onQuotation = True
                        Words(j) = "\cf5" & Chr(32) & Words(j)
                        
                        GoTo EndIt
                    Else
                        onQuotation = False
                        Words(j) = Words(j) & "\cf0"
                        
                        GoTo EndIt
                    End If
                End If
            'Alles andere
            Case Else
                'Wörter von Delimiter trennen
                Pos = InStr(1, Delimiter, Right(Words(j), 1))
                
                If Pos > 0 Then
                    Word = Delimiters(Pos - 1)
                    If Len(Words(j)) <= 0 Then GoTo EndIt
                    Words(j) = Left(Words(j), Len(Words(j)) - Len(Word))
                End If
                
                'Nur einfärben wenn nicht in Kommentar
                If onComment = False And onSingleComment = False And onQuotation = False Then
                    'Keywords
                    If InStr(1, Keywords, Chr(32) & Words(j) & Chr(32), vbTextCompare) > 0 Then
                        If ProperCase > 0 Then
                            Words(j) = "\cf2" & onBold & StrConv(Words(j), CInt(ProperCase)) & offBold & "\cf0 "
                        Else
                            Words(j) = "\cf2" & onBold & Words(j) & offBold & "\cf0 "
                        End If
                    End If
                        
                    'KeywordsAlt
                    If InStr(1, KeywordsAlt, Chr(32) & Words(j) & Chr(32), vbTextCompare) > 0 Then
                        If ProperCase > 0 Then
                            Words(j) = "\cf3" & onBold & StrConv(Words(j), CInt(ProperCase)) & offBold & "\cf0 "
                        Else
                            Words(j) = "\cf3" & onBold & Words(j) & offBold & "\cf0 "
                        End If
                    End If
                End If

                If Pos > 0 Then
                    'Operatoren einfärben
                    If onComment = False And onSingleComment = False And onQuotation = False Then
                        If InStr(1, Operator, Word, vbTextCompare) > 0 Then
                            Word = "\cf7" & Word & "\cf0"
                        End If
                    End If
                    
                    'Kommenatare und Strings
                    Select Case Word
                        'Einfacher Kommentar
                        Case "¨"
                            If onQuotation = False Then
                                If onSingleComment = False Then
                                    onSingleComment = True
                                    Word = "\cf4" & Chr(32) & Word
                                    
                                    GoTo EndColor
                                End If
                            End If
                        'Kommentar Anfang
                        Case "´"
                            If onQuotation = False Then
                                If onComment = False Then
                                    onComment = True
                                    Word = "\cf4" & Chr(32) & Word
                                End If
                            End If
                        'Kommentar Ende
                        Case "`"
                            If onQuotation = False Then
                                If onComment = True Then
                                    onComment = False
                                    Word = Word & "\cf0"
                                End If
                            End If
                        'Strings
                        Case Chr(34)
                            If onComment = False And onSingleComment = False Then
                                If onQuotation = False Then
                                    onQuotation = True
                                    Word = "\cf5" & Chr(32) & Word
                                    
                                    GoTo EndColor
                                Else
                                    onQuotation = False
                                    Word = Word & "\cf0"
                                
                                    GoTo EndColor
                                End If
                            End If
                    End Select
EndColor:
                    Words(j) = Words(j) & Word
                End If
        End Select
EndIt:
    Next j
EndLine:
    
    'Wörter zusammensetzen
    Line = Join(Words, Chr(32))
      
    'Ersetzet Kommentarzeichen
    Line = Replace(Line, "¨", CommentChar, , , vbTextCompare)
    Line = Replace(Line, "´", CommentStartChar, , , vbTextCompare)
    Line = Replace(Line, "`", CommentEndChar, , , vbTextCompare)

    'Leerzeichen nach Delimiters ersetzen
    For j = 0 To UBound(Delimiters)
        Line = Replace(Line, Delimiters(j) & Chr(32), Delimiters(j))
    Next j
    
    'Einfache Kommentare ausschalten
    If onSingleComment = True Then
        Line = Line & "\cf0"
    End If
    
    Lines(i) = Line
Quit:
Next i

'Alle Zeilen zusammenfügen
RTF = Join(Lines, vbCrLf & "\par ")
'Header erschaffen
Header = CreateHeader(FontName, FontSize)

'RTF zurückgeben
ALLRTF = Header & RTF & vbCrLf & "}"

frmMain.prgBar.Value = 0

ColorCode = ALLRTF
End Function

Public Function ColorHTML(Text As String, FontName As String, FontSize As Integer) As String
Dim Lines() As String
Dim Line As String
Dim Words() As String
Dim Pos As Integer
Dim Word As String

Dim Operators() As String

Dim RTF As String
Dim ALLRTF As String
Dim Header As String

Dim onComment As Boolean
Dim onScript As Boolean
Dim onQuotation As Boolean
Dim onTag As Boolean

Dim i As Integer
Dim j As Integer
Dim x As Integer

Lines = Split(Text, vbCrLf)

Operators = Split(Operator, Chr(32))

For i = 0 To UBound(Lines)
    DoEvents
    If Len(Text) > 40000 Then frmMain.prgBar.Value = i / (UBound(Lines) / 100)
    
    onQuotation = False
    
    Line = Lines(i)

    Line = Replace(Line, "\", "\\")
    Line = Replace(Line, "}", "\}")
    Line = Replace(Line, "{", "\{")
    
    Line = Replace(Line, "<!--", "´", , , vbTextCompare)
    Line = Replace(Line, "-->", "`", , , vbTextCompare)
    Line = Replace(Line, "<script", "¨", , , vbTextCompare)
    Line = Replace(Line, "</script>", "¯", , , vbTextCompare)
    
    For j = 0 To UBound(Delimiters)
        Line = Replace(Line, Delimiters(j), Delimiters(j) & Chr(32))
    Next j
    
    Words = Split(Line)

    For j = 0 To UBound(Words)
         Select Case UCase(Words(j))
            Case "¨"
                If onQuotation = False Then
                    If onComment = False Then
                        onScript = True
                        onComment = True
                        Words(j) = "\cf6" & Chr(32) & Words(j)
                    End If
                End If
            Case "¯"
                If onQuotation = False Then
                    If onComment = True Then
                        onScript = False
                        onComment = False
                        Words(j) = Words(j) & "\cf0"
                    End If
                End If
            Case "´"
                If onQuotation = False Then
                    If onComment = False Then
                        onComment = True
                        Words(j) = "\cf4" & Chr(32) & Words(j)
                    End If
                End If
            Case "`"
                If onQuotation = False Then
                    If onComment = True Then
                        If onScript = False Then
                            onComment = False
                            Words(j) = Words(j) & "\cf0"
                        End If
                    End If
                End If
            Case "<"
               If onComment = False And onQuotation = False Then
                    onTag = True
                    Words(j) = "\cf2" & Chr(32) & Words(j)
                
                    GoTo EndIt
                End If
            Case ">"
                If onComment = False And onQuotation = False Then
                    onTag = False
                    Words(j) = Words(j) & "\cf0"
                        
                    GoTo EndIt
                End If
            Case Chr(34)
                If onComment = False Then
                    If onTag = True Then
                        If onQuotation = False Then
                            onQuotation = True
                            Words(j) = "\cf5" & Chr(32) & Words(j)
                                
                            GoTo EndIt
                        Else
                            onQuotation = False
                            Words(j) = Words(j) & "\cf2"
                                
                            GoTo EndIt
                        End If
                    End If
                End If
            Case Else
                Pos = InStr(1, Delimiter, Right(Words(j), 1))
                
                If Pos > 0 Then
                    Word = Delimiters(Pos - 1)
                    If Len(Words(j)) <= 0 Then GoTo EndIt
                    Words(j) = Left(Words(j), Len(Words(j)) - Len(Word))
                End If

                If Pos > 0 Then
                    Select Case Word
                        Case "¨"
                            If onQuotation = False Then
                                If onComment = False Then
                                    onScript = True
                                    onComment = True
                                    Word = "\cf6" & Chr(32) & Word
                                End If
                            End If
                        Case "¯"
                            If onQuotation = False Then
                                If onComment = True Then
                                    onScript = False
                                    onComment = False
                                    Word = Word & "\cf0"
                                End If
                            End If
                        Case "´"
                            If onQuotation = False Then
                                If onComment = False Then
                                    onComment = True
                                    Word = "\cf4" & Chr(32) & Word
                                End If
                            End If
                        Case "`"
                            If onQuotation = False Then
                                If onComment = True Then
                                    If onScript = False Then
                                        onComment = False
                                        Word = Word & "\cf0"
                                    End If
                                End If
                            End If
                        Case "<"
                            If onComment = False And onQuotation = False Then
                                onTag = True
                                Word = "\cf2" & Chr(32) & Word
                               
                                GoTo EndColor
                            End If
                        Case ">"
                            If onComment = False And onQuotation = False Then
                                onTag = False
                                Word = Word & "\cf0"
                                
                                GoTo EndColor
                            End If
                        Case Chr(34)
                            If onComment = False Then
                                If onTag = True Then
                                    If onQuotation = False Then
                                        onQuotation = True
                                        Word = "\cf5" & Chr(32) & Word
                                            
                                        GoTo EndColor
                                    Else
                                        onQuotation = False
                                        Word = Word & "\cf2"
                                        
                                        GoTo EndColor
                                    End If
                                End If
                            End If
                    End Select
EndColor:
                    Words(j) = Words(j) & Word
                End If
        End Select
EndIt:
    Next j
EndLine:
    
    Line = Join(Words, Chr(32))
    
    'Leerzeichen nach Delimiters ersetzen
    For j = 0 To UBound(Delimiters)
        Line = Replace(Line, Delimiters(j) & Chr(32), Delimiters(j))
    Next j

    Line = Replace(Line, "´", "<!--", , , vbTextCompare)
    Line = Replace(Line, "`", "-->", , , vbTextCompare)
    Line = Replace(Line, "¨", "<script", , , vbTextCompare)
    Line = Replace(Line, "¯", "</script>", , , vbTextCompare)
    
    Lines(i) = Line
    
Quit:
Next i

RTF = Join(Lines, vbCrLf & "\par ")
Header = CreateHeader(FontName, FontSize)

ALLRTF = Header & RTF & vbCrLf & "}"

frmMain.prgBar.Value = 0

ColorHTML = ALLRTF
End Function

Public Function ColorText(Text As String, FontName As String, FontSize As Integer) As String
Dim Lines() As String

Dim Header As String
Dim ALLRTF As String

Header = CreateHeader(FontName, FontSize)

Lines = Split(Text, vbCrLf)
Text = Join(Lines, vbCrLf & "\par ")

ALLRTF = Header & Text & vbCrLf & "}"

ColorText = ALLRTF
End Function

Private Function CreateHeader(FontName As String, FontSize As Integer) As String
Dim i As Integer
Dim H1 As String
Dim H2 As String
Dim ColorH As String

'Erstelle Colortable
ColorH = "{\colortbl " & ConvertColorToRTF(0)

For i = 1 To 2
    ColorH = ColorH & ConvertColorToRTF(KeywordColor(i))
Next i

ColorH = ColorH & ConvertColorToRTF(CommentColor)
ColorH = ColorH & ConvertColorToRTF(QuotationColor)
ColorH = ColorH & ConvertColorToRTF(ScriptColor)
ColorH = ColorH & ConvertColorToRTF(OperatorColor)
ColorH = ColorH & ";}"

'Header
H1 = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 " & FontName & ";}}"
H2 = "\viewkind4\uc1\pard\f0\fs" & Round(FontSize * 2) & " "

'Zusammenfügen
CreateHeader = H1 & vbCrLf & ColorH & vbCrLf & H2
End Function

Private Function ConvertColorToRTF(LongColor As Long) As String
Dim ColorRTFCode As String
Dim LC As Long
    
LC = LongColor And &H10000FF
ColorRTFCode = ";\red" & LC
LC = (LongColor And &H100FF00) / (2 ^ 8)
ColorRTFCode = ColorRTFCode & "\green" & LC
LC = (LongColor And &H1FF0000) / (2 ^ 16)
ColorRTFCode = ColorRTFCode & "\blue" & LC
ColorRTFCode = ColorRTFCode & ""
    
'Return Var
ConvertColorToRTF = ColorRTFCode
End Function

'INI laden
Public Function LoadIni(Filename As String, KeySection As String, KeyKey As String) As String
Dim Result As Long
Dim Entry As String * 10000

Result = GetPrivateProfileString(KeySection, KeyKey, Filename, Entry, Len(Entry), Filename)

'Wenn fehlgeschlagen, dann String leer ausgeben
If Right(TrimWS(Entry), 3) = "txt" Or Right(TrimWS(Entry), 3) = "ini" Or Right(TrimWS(Entry), 3) = "stx" Then
    Entry = ""
End If

If Result <> 0 Then
    LoadIni = TrimWS(Entry)
End If
End Function

'String trimmen
Public Function TrimWS(Text As String) As String
Dim LeftEnd As Long
Dim RightEnd As Long

'Start suchen
For LeftEnd = 1 To Len(Text)
    If Asc(Mid(Text, LeftEnd, 1)) > vbKeySpace Then Exit For
Next LeftEnd

'Ende suchen
For RightEnd = Len(Text) To LeftEnd + 1 Step -1
    If Asc(Mid(Text, RightEnd, 1)) > vbKeySpace Then Exit For
Next RightEnd

'Fertig
TrimWS = Mid(Text, LeftEnd, RightEnd - LeftEnd + 1)
End Function
