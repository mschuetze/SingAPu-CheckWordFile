' version 0.21.15

'----------------------------------------------------------
'----- SET GLOBAL VARIABLES -----
'----------------------------------------------------------
Option Explicit
Public DebugMode As Boolean
Public NameOfFormat As String
Public NameOfFormatAfter As String
Public NameOfFormatBefore As String
Dim multiStyles As String, I As Long
Dim aStyleList As Variant
Dim counter As Long, s As String
Dim correctFormat As Boolean
Dim logFileName As String
Dim NameContainsSpecialChars As Boolean
Dim char As String
Dim logFilePath As String
Dim logFile As Long
Dim para As Paragraph
Dim paraText As String
Dim first40Chars As String
Dim previousPara As Paragraph
Dim nextPara As Paragraph
Dim styleCount As Long
Dim logEntries As Collection ' Declare a global collection for log entries





Sub InitializeLog()
    ' Initialize the log entries collection
    Set logEntries = New Collection
End Sub

Sub AddLogEntry(entry As String)
    ' Add a log entry to the collection
    logEntries.Add Now & " : " & entry
End Sub

Sub WriteLogEntries()
    Dim logFileNumber As Long
    Dim entry As Variant

    ' Get a free file number
    logFileNumber = FreeFile

    ' Open the log file for appending
    Open logFilePath For Append As #logFileNumber

    ' Write each log entry to the file
    For Each entry In logEntries
        Print #logFileNumber, entry & vbCrLf & "----"
    Next entry

    ' Close the file
    Close #logFileNumber
End Sub





Sub SingAPu_CheckWordFile()
    InitializeLog ' Initialize the logEntries collection
    DebugMode = False ' Enable debug messages
    
    Application.ScreenUpdating = False  ' GUI-Belastung reduzieren

'----------------------------------------------------------
'----- CHECK FILE NAME -----
'----------------------------------------------------------

    If DebugMode Then MsgBox "Launching: CHECK FILE NAME"
    Dim fileName As String
    Dim baseFileName As String
    Dim invalidChars As String
    Dim i As Long
    Dim currentChar As String
    Dim umlautChars As String
    Dim emptySpaceChar As String
    Dim foundInvalid As Boolean
    Dim invalidList As String
    ' Missing declarations (required with Option Explicit)
    Dim cleanedFileName As String
    Dim fileToCheckPath As String
    
    fileName = ActiveDocument.Name
    ' Entfernen der Dateiendung (alles nach dem letzten Punkt)
    If InStrRev(fileName, ".") > 0 Then
        baseFileName = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        baseFileName = fileName
    End If
    ' MsgBox "baseFileName: " & baseFileName

    ' NEW: Prüfen, ob die Dateiendung .docx ist; falls nicht, ins Log schreiben
    Dim fileExt As String
    If InStrRev(fileName, ".") > 0 Then
        fileExt = LCase(Mid(fileName, InStrRev(fileName, ".") + 1))
    Else
        fileExt = ""
    End If

    If fileExt <> "docx" Then
        AddLogEntry "Dateiendung ist nicht .docx: " & fileName
    End If

    invalidChars = "!@#$%^&*()+={}[]|\:;""'<>,.?/~`" ' Hier definierst du die Sonderzeichen, die du überprüfen möchtest.
    umlautChars = "äöüÄÖÜß" ' Umlaute und Sonderzeichen, die überprüft werden sollen.
    emptySpaceChar = " " ' Leerzeichen wird separat aufgeführt, um eine aussagekräftige Fehlermeldung zu generieren
    
    foundInvalid = False
    invalidList = "" ' Leere Liste für ungültige Zeichen
    
    ' Überprüfen des Dateinamens auf unerlaubte Sonderzeichen
    For i = 1 To Len(baseFileName)
        currentChar = Mid(baseFileName, i, 1)
        ' MsgBox "currentChar: " & currentChar
        
        ' Überprüfen auf Sonderzeichen / Umlaute / Leerzeichen und Liste aufbauen
        If InStr(invalidChars, currentChar) > 0 Then
            If InStr(invalidList, currentChar) = 0 Then invalidList = invalidList & currentChar & " "
            foundInvalid = True
        End If
        If InStr(umlautChars, currentChar) > 0 Then
            If InStr(invalidList, currentChar) = 0 Then invalidList = invalidList & currentChar & " "
            foundInvalid = True
        End If
        If InStr(emptySpaceChar, currentChar) > 0 Then
            If InStr(invalidList, "Leerzeichen") = 0 Then invalidList = invalidList & "Leerzeichen "
            foundInvalid = True
        End If
    Next i
    
    ' Falls keine ungültigen Zeichen gefunden wurden
    If foundInvalid Then
        MsgBox "Die folgenden Sonderzeichen wurden im Dateinamen gefunden und müssen zunächst ersetzt werden: " & vbCrLf & invalidList, vbExclamation
        Exit Sub
    End If
    If DebugMode Then MsgBox "Done: CHECK FILE NAME"





'----------------------------------------------------------
'----- DELETE LOG FILE, IF EXISTS -----
'----------------------------------------------------------
If DebugMode Then MsgBox "Launching: DELETE LOG FILE"

' Set name of log file
logFileName = "log" & ".txt"
 
' Set the path for the log file
logFilePath = ActiveDocument.Path & "/" & logFileName
' MsgBox "logFilePath: " & logFilePath

If Dir(logFilePath) <> "" Then
    ' MsgBox "Log-Datei besteht bereits. Wird gelöscht."
    Kill logFilePath
End If

If DebugMode Then MsgBox "Done: DELETE LOG FILE"





'----------------------------------------------------------
'----- CHECK IF FIRST PARAGRAPH HAS A PIPE IN IT, IF SO CHECK IF PARAGRAPH IS FORMAT X, IF NOT SET THE CORRECT FORMAT -----
'----------------------------------------------------------

If DebugMode Then MsgBox "Launching: CHECK FOR PIPE"

Dim firstParagraph As Paragraph
Dim appliedStyle As String

' Den ersten Absatz abrufen
Set firstParagraph = ActiveDocument.Paragraphs(1)

' Den Inhalt des ersten Absatzes abrufen
paraText = firstParagraph.Range.Text

' Prüfen, ob der Text das Zeichen "|" enthält
If InStr(paraText, "|") > 0 Then
    ' Den Namen der angewendeten Formatvorlage abrufen
    appliedStyle = firstParagraph.Style
    ' Prüfen, ob das Absatzformat "SuS_Mengentext" angewendet wurde
    If appliedStyle <> "SuS_Mengentext" Then
        ' MsgBox "Der erste Absatz enthält zwar das Zeichen '|', hat aber nicht das Format 'SuS_Mengentext'.", vbExclamation
        ' Absatzformat zuweisen
        firstParagraph.Style = "SuS_Mengentext"
    End If
Else
    ' MsgBox "Der erste Absatz enthält das Zeichen '|' NICHT.", vbExclamation
    logFile = FreeFile
    Open logFilePath For Append As logFile
    Print #logFile, Now & vbCrLf & "Im ersten Absatz fehlt das Zeichen '|' (Pipe)." & vbCrLf & "----" & vbCrLf
    Close logFile
End If

If DebugMode Then MsgBox "Done: CHECK FOR PIPE"





'----------------------------------------------------------
'----- SET HEADER FORMATS -----
'----------------------------------------------------------

If DebugMode Then MsgBox "Launching: SET HEADER FORMATS"

' SET FORMAT OF FIRST PARAGRAPH
' ActiveDocument.Paragraphs.First.Range.Select
' search_firstPara
'Format will be set in Sub if check is TRUE

' SET FORMAT OF SECOND PARAGRAPH
ActiveDocument.Paragraphs(2).Style = "SuS_Headline"

' SET FORMAT OF THIRD PARAGRAPH
ActiveDocument.Paragraphs(3).Style = "SuS_Subhead1"

If DebugMode Then MsgBox "Done: SET HEADER FORMATS"





'----------------------------------------------------------
'----- CHECK NUMBER OF FORMAT INSTANCES  -----
'----------------------------------------------------------

' Style counts are now performed in ConsolidatedParagraphChecks





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER NEXT PARAGRAPH IS FORMAT Y -----
'----------------------------------------------------------
    
Dim IsFound As Boolean

' Diese Checks sind jetzt in ConsolidatedParagraphChecks integriert
' NameOfFormat = "SuS_Subhead2"
' multiStyles = "SuS_Autorname"
' IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

' NameOfFormat = "SuS_Bild/Tabellenunterschrift"
' multiStyles = "SuS_Mengentext,SuS_Kastentext,SuS_Absatzheadline,SuS_Unter_Absatzheadline,SuS_Kasten_Absatzheadline"
' IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

' NEW: Überprüfe zusätzlich die bedingte Regel für "SuS_Bilddateiname"
' SuS_Kastentext darf folgen, wenn der Absatz vor SuS_Bilddateiname eine SuS_Kastenheadline mit "Porträt" ist
' Sonst darf nur SuS_Bild/Tabellenunterschrift oder SuS_Autor_Kurzbiografie folgen
' check_bilddateiname_follows_rule

' NameOfFormat = "SuS_Links_und_Literatur_Headline"
' multiStyles = "SuS_Links_und_Literatur_Text"
' IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

' NameOfFormat = "SuS_Links_und_Literatur_Text"
' multiStyles = "SuS_Links_und_Literatur_Text"
' IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER NEXT PARAGRAPH IS NOT FORMAT Y -----
'----------------------------------------------------------

' NameOfFormat = "SuS_Kastenheadline"
' multiStyles = "SuS_Kastenheadline"
' IsFound = FindParagraphAfterMustNotBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER PREVIOUS PARAGRAPH IS FORMAT Y -----
'----------------------------------------------------------

' ENTFERNT: NameOfFormat = "SuS_Bilddateiname" wird jetzt in check_bilddateiname_follows_rule gehandhabt





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER PREVIOUS PARAGRAPH IS NOT FORMAT Y -----
'----------------------------------------------------------

' NameOfFormat = "SuS_Kastenheadline"
' multiStyles = "SuS_Kastenheadline"
' IsFound = FindParagraphBeforeMustNotBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER IT CONTAINS SPECIAL CHARACTERS -----
'----------------------------------------------------------

If DebugMode Then MsgBox "Start: 'find special chars in image reference'"

Dim Absatz As Paragraph
Dim Formatvorlage As String
Dim baseAbsatzText As String
Dim j As Long
Dim dotPos As Long
Dim fileExtension As String
Dim valid As Boolean
Dim Dateiendung As Variant
Dim EndungGefunden As Boolean
Dim ext As Variant ' loop variable for extensions

' Die Formatvorlage, die du suchen möchtest
Formatvorlage = "SuS_Bilddateiname"

' Erlaubte Bild-Endungen definieren
Dateiendung = Array(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg")

' Gehe alle Absätze im Dokument durch
For Each Absatz In ActiveDocument.Paragraphs
    ' Wenn der Absatz die angegebene Formatvorlage hat
    If Absatz.Style = Formatvorlage Then
        foundInvalid = False
        invalidList = "" ' Leere Liste für ungültige Zeichen

        ' Hole und bereinige den Text des Absatzes
        paraText = Absatz.Range.Text
        cleanedFileName = Replace(paraText, Chr(13), "")
        cleanedFileName = Replace(cleanedFileName, Chr(10), "")
        cleanedFileName = Trim(cleanedFileName)

        ' Falls ein kompletter Pfad angegeben ist, nur den Dateinamen extrahieren
        If InStrRev(cleanedFileName, "/") > 0 Then cleanedFileName = Mid(cleanedFileName, InStrRev(cleanedFileName, "/") + 1)
        If InStrRev(cleanedFileName, "\") > 0 Then cleanedFileName = Mid(cleanedFileName, InStrRev(cleanedFileName, "\") + 1)

        ' Optional: für die Zeichenprüfung die Dateiendung entfernen (falls vorhanden)
        If InStrRev(cleanedFileName, ".") > 0 Then
            baseAbsatzText = Left(cleanedFileName, InStrRev(cleanedFileName, ".") - 1)
        Else
            baseAbsatzText = cleanedFileName
        End If

        ' Prüfen, ob die Dateiendung vorhanden und erlaubt ist
        EndungGefunden = False
        dotPos = InStrRev(cleanedFileName, ".")
        If dotPos > 0 Then
            fileExtension = LCase(Mid(cleanedFileName, dotPos)) ' inkl. führendem Punkt
            For Each ext In Dateiendung
                If fileExtension = LCase(ext) Then
                    EndungGefunden = True
                    Exit For
                End If
            Next ext
        Else
            fileExtension = ""
        End If

        If Not EndungGefunden Then
            AddLogEntry "Bildverweis hat fehlende oder ungültige Dateiendung: " & cleanedFileName & " (erwartet: " & Join(Dateiendung, ", ") & ")"
        End If

        ' Überprüfe jeden Charakter im (ohne Endung) Dateinamen auf Sonderzeichen / Umlaute / Leerzeichen
        For j = 1 To Len(baseAbsatzText)
            currentChar = Mid(baseAbsatzText, j, 1)
            If InStr(invalidChars, currentChar) > 0 Then
                If InStr(invalidList, currentChar) = 0 Then invalidList = invalidList & currentChar & " "
                foundInvalid = True
            End If
            If InStr(umlautChars, currentChar) > 0 Then
                If InStr(invalidList, currentChar) = 0 Then invalidList = invalidList & currentChar & " "
                foundInvalid = True
            End If
            If InStr(emptySpaceChar, currentChar) > 0 Then
                If InStr(invalidList, currentChar) = 0 Then invalidList = invalidList & "Leerzeichen "
                foundInvalid = True
            End If
        Next j

        If foundInvalid Then
            AddLogEntry "Bildverweis " & cleanedFileName & " enthält folgende Sonderzeichen: " & invalidList
        End If

        ' Prüfen, ob die referenzierte Bilddatei im Dokument-Ordner existiert
        ' Wenn Dokument nicht gespeichert, kann nicht geprüft werden
        If ActiveDocument.Path = "" Then
            AddLogEntry "Dokument nicht gespeichert: Kann nicht prüfen, ob Bilddatei vorhanden ist: " & cleanedFileName
            GoTo NextAbsatz
        End If

        fileToCheckPath = ActiveDocument.Path & "/" & cleanedFileName
        If Dir(fileToCheckPath) <> "" Then
            ' exakte Datei gefunden -> alles gut
        Else
            ' exakte Datei nicht gefunden -> prüfen, ob eine Datei mit gleichem Basisnamen und einer erlaubten Endung existiert
            Dim foundAlternative As Boolean
            Dim baseNameOnly As String
            Dim altPath As String
            Dim matchedFile As String
            foundAlternative = False
            dotPos = InStrRev(cleanedFileName, ".")
            If dotPos > 0 Then
                baseNameOnly = Left(cleanedFileName, dotPos - 1)
            Else
                baseNameOnly = cleanedFileName
            End If
            For Each ext In Dateiendung
                altPath = ActiveDocument.Path & "/" & baseNameOnly & ext
                If Dir(altPath) <> "" Then
                    foundAlternative = True
                    matchedFile = baseNameOnly & ext
                    Exit For
                End If
            Next ext

            If Not foundAlternative Then
                AddLogEntry "Bilddatei nicht gefunden: " & cleanedFileName & " (verweisender Absatz: [" & First40Characters(Absatz) & "])"
            Else
                ' Alternative mit anderer Endung gefunden -> OK, kein Log-Eintrag nötig
            End If
        End If
    End If
NextAbsatz:
Next Absatz

'----------------------------------------------------------
'----- CHECK IF NUMBER OF INSTANCES FORMAT X IS AN INTEGER MULTIPLE OF 2 (FORMAT ALWAYS NEEDS TO BE CLOSED) -----
'----------------------------------------------------------

' Style count for SuS_Kastenheadline is now performed in ConsolidatedParagraphChecks





'----------------------------------------------------------
'----- CHECK IF PREVIOUS PARAGRAPH OF ODD OCCURENCES OF FORMAT X IS FORMAT Y -----
'----------------------------------------------------------

' Diese Checks sind jetzt in ConsolidatedParagraphChecks integriert
' CHECK FOR SuS_Kastenheadline
' NameOfFormat = "SuS_Kastenheadline"
' multiStyles = "SuS_Mengentext"
' check_style_before_odd





'----------------------------------------------------------
'----- CHECK IF ODD OCCURENCES OF FORMAT X IS NOT EMPTY ---
'----------------------------------------------------------

' Diese Checks sind jetzt in ConsolidatedParagraphChecks integriert
' NameOfFormat = "SuS_Kastenheadline"
' check_odd_kastenheadline_empty

' NEW: Prüfe bei ungeraden SuS_Kastenheadline mit "Listing", dass alle folgenden Absätze bis zur nächsten SuS_Kastenheadline SuS_Quellcode sind
' check_listing_followed_by_quellcode

'----------------------------------------------------------
'----- CHECK IF ONLY STYLES WITH STRING 'SuS_' ARE BEING USED ---
'----------------------------------------------------------
' check_invalid_styles

' Konsolidierte Paragraph-Checks durchführen
ConsolidatedParagraphChecks

'----------------------------------------------------------
'----- CHECK IF CERTAIN FORMATS ARE NOT ITALIC ---
'----------------------------------------------------------
check_italic_formats



'----------------------------------------------------------
'----- CHECK FOR DIRECT BOLD/ITALIC WITHOUT CHARACTER STYLE ---
'----------------------------------------------------------
check_direct_bold_italic_without_character_style


'----------------------------------------------------------
'----- WRITE ALL LOG ENTRIES TO THE FILE ------------------
'----------------------------------------------------------

    WriteLogEntries





'----------------------------------------------------------
'----- END OF SCRIPT MESSAGE -----
'----------------------------------------------------------

Application.ScreenUpdating = True   ' Bildschirmaktualisierung wieder aktivieren
MsgBox "Ich habe fertig."

End Sub




















Function First40Characters(para As Paragraph) As String
    Dim tmpText As String
    
    ' Den Text holen und Steuerzeichen sofort entfernen
    tmpText = para.Range.Text
    tmpText = Replace(tmpText, vbCr, "")
    tmpText = Replace(tmpText, vbLf, "")
    
    ' Erst danach auf 40 Zeichen kürzen und trimmen
    First40Characters = Trim(Left$(tmpText, 40))
End Function







Sub UpdateCount(styleNames() As String, styleCounts() As Long, ByRef count As Long, ByVal key As String)
    Dim i As Long
    For i = 0 To count - 1
        If styleNames(i) = key Then
            styleCounts(i) = styleCounts(i) + 1
            Exit Sub
        End If
    Next i
    ' Neu hinzufügen
    If count = 0 Then
        ' Array ist bereits ReDim(0), setze direkt
        styleNames(0) = key
        styleCounts(0) = 1
    Else
        ReDim Preserve styleNames(count)
        ReDim Preserve styleCounts(count)
        styleNames(count) = key
        styleCounts(count) = 1
    End If
    count = count + 1
End Sub







Sub check_italic_formats()
    '-------------------------------------------
    ' Überprüfe, ob bestimmte Formate kursiv sind
    ' Formate dürfen weder als Formatvorlage noch als Standardformatierung kursiv sein
    '-------------------------------------------
    If DebugMode Then MsgBox "Start: Sub 'check_italic_formats'"

    Dim searchRange As Range
    Dim paraStyle As String
    Dim prohibitedFormats As Variant
    Dim fmt As Variant
    Dim foundText As String

    prohibitedFormats = Array( _
        "SuS_Headline", "SuS_Subhead1", "SuS_Bild/Tabellenunterschrift", _
        "SuS_Absatzheadlines", "SuS_Kastenheadline", "SuS_Tabellenkopf", _
        "SuS_Quellcode", "SuS_Quellcode_Kommentar", _
        "SuS_Links_und_Literatur_Headline", "SuS_Links_und_Literatur_Text")

    Application.ScreenUpdating = False

    ' Einmal über das gesamte Dokument nach direkt kursivem Text suchen
    Set searchRange = ActiveDocument.Content
    With searchRange.Find
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While searchRange.Find.Execute
        ' Nur Absatzformate prüfen, keine Zeichenformate
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            paraStyle = searchRange.Paragraphs(1).Style.NameLocal
            For Each fmt In prohibitedFormats
                If paraStyle = fmt Then
                    foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
                    AddLogEntry "Format '" & paraStyle & "' darf nicht kursiv sein: [" & First40Characters(searchRange.Paragraphs(1)) & "] – Textstelle: [" & foundText & "]"
                    Exit For
                End If
            Next fmt
        End If
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    Application.ScreenUpdating = True

    If DebugMode Then MsgBox "Ende: Sub 'check_italic_formats'"
End Sub




Sub check_direct_bold_italic_without_character_style()
    '-------------------------------------------
    ' Prüft auf direktes Fett/Kursiv im Dokument ohne zugeordnetes Zeichenformat
    ' Zusätzlich wird geprüft, ob der Absatzstil für das gefundene Format erlaubt ist.
    '-------------------------------------------
    If DebugMode Then MsgBox "Start: Sub 'check_direct_bold_italic_without_character_style'"

    Dim searchRange As Range
    Dim first40Chars As String
    Dim foundText As String
    Dim paraStyle As String
    Dim allowedBoldFormats As Variant
    Dim allowedItalicFormats As Variant
    Dim fmt As Variant
    Dim styleAllowed As Boolean

    allowedBoldFormats = Array("SuS_Autorname", "SuS_Kastenheadline", "SuS_Absatzheadline", "SuS_Unter_Absatzheadline", "SuS_Kasten_Absatzheadline", "SuS_Tabellenkopf", "SuS_Links_und_Literatur_Headline", "SuS_Interview_Frage", "SuS_Interview_Zitat")
    allowedItalicFormats = Array("SuS_Autor_Kurzbiografie", "SuS_Interview_Zitat")

    ' Screening-Optimierung: Screen-Refresh deaktivieren
    Application.ScreenUpdating = False

    ' -------------------------------------------------------
    ' Hilfsfunktion: führt eine Format-Suche durch und loggt
    ' -------------------------------------------------------

    ' 1) Bold
    Set searchRange = ActiveDocument.Content
    searchRange.Find.ClearFormatting
    With searchRange.Find
        .Font.Bold = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop   ' bleibt Stop – aber Range-Verwaltung wird korrigiert
    End With

    Do While searchRange.Find.Execute
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            paraStyle = searchRange.Paragraphs(1).Style.NameLocal
            styleAllowed = False
            For Each fmt In allowedBoldFormats
                If paraStyle = fmt Then styleAllowed = True: Exit For
            Next fmt
            If Not styleAllowed Then
                foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
                AddLogEntry "Direktes Fett ohne Zeichenformat in [" & First40Characters(searchRange.Paragraphs(1)) & "]: [" & foundText & "]"
            End If
        End If
        ' Korrekte Weiterbewegung: Collapse ans Ende statt manuellen Range-Reset
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    ' 2) Italic
    Set searchRange = ActiveDocument.Content
    searchRange.Find.ClearFormatting
    With searchRange.Find
        .Font.Italic = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While searchRange.Find.Execute
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            paraStyle = searchRange.Paragraphs(1).Style.NameLocal
            styleAllowed = False
            For Each fmt In allowedItalicFormats
                If paraStyle = fmt Then styleAllowed = True: Exit For
            Next fmt
            If Not styleAllowed Then
                foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
                AddLogEntry "Direktes Kursiv ohne Zeichenformat in [" & First40Characters(searchRange.Paragraphs(1)) & "]: [" & foundText & "]"
            End If
        End If
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    ' 3) SmallCaps
    Set searchRange = ActiveDocument.Content
    searchRange.Find.ClearFormatting
    With searchRange.Find
        .Font.SmallCaps = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While searchRange.Find.Execute
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
            AddLogEntry "Direkte Kapitälchen ohne Zeichenformat in [" & First40Characters(searchRange.Paragraphs(1)) & "]: [" & foundText & "]"
        End If
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    ' 4) Superscript
    Set searchRange = ActiveDocument.Content
    searchRange.Find.ClearFormatting
    With searchRange.Find
        .Font.Superscript = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While searchRange.Find.Execute
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
            AddLogEntry "Direkte Hochstellung ohne Zeichenformat in [" & First40Characters(searchRange.Paragraphs(1)) & "]: [" & foundText & "]"
        End If
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    ' 5) Subscript
    Set searchRange = ActiveDocument.Content
    searchRange.Find.ClearFormatting
    With searchRange.Find
        .Font.Subscript = True
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While searchRange.Find.Execute
        If searchRange.Style.Type <> wdStyleTypeCharacter Then
            foundText = Trim(Replace(Replace(searchRange.Text, Chr(13), ""), Chr(10), ""))
            AddLogEntry "Direkte Tiefstellung ohne Zeichenformat in [" & First40Characters(searchRange.Paragraphs(1)) & "]: [" & foundText & "]"
        End If
        searchRange.Collapse wdCollapseEnd
        searchRange.End = ActiveDocument.Content.End
    Loop

    Application.ScreenUpdating = True

    If DebugMode Then MsgBox "Ende: Sub 'check_direct_bold_italic_without_character_style'"
End Sub










Sub ConsolidatedParagraphChecks()
    If DebugMode Then MsgBox "Start: ConsolidatedParagraphChecks"

    Dim para As Paragraph
    Dim styleCount As Long
    Dim invalidStyles As Collection
    Dim styleName As Variant
    Dim wrtstring As String
    Dim st As Style
    Dim searchRange As Range
    Dim prevPara As Paragraph
    Dim nextPara As Paragraph
    Dim headlineText As String
    Dim first40Chars As String
    Dim isValidCondition As Boolean
    Dim previousPara As Paragraph
    Dim correctFormat As Boolean
    Dim aStyleList As Variant
    Dim counter As Long
    Dim paraText As String
    Dim p As Paragraph
    Dim offendingText As String
    Dim styleNames() As String
    Dim styleCounts() As Long
    Dim styleCountIndex As Long

    Set invalidStyles = New Collection
    ReDim styleNames(0)
    ReDim styleCounts(0)
    styleCountIndex = 0
    styleCount = 0

    For Each para In ActiveDocument.Paragraphs
        ' check_invalid_styles für Paragraph-Styles
        styleName = para.Style.NameLocal
        If Left(styleName, 4) <> "SuS_" Then
            On Error Resume Next
            invalidStyles.Add styleName, styleName
            On Error GoTo 0
        End If

        ' Count all styles
        UpdateCount styleNames, styleCounts, styleCountIndex, styleName

        ' check_bilddateiname_follows_rule
        If para.Style = "SuS_Bilddateiname" Then
            ' Überprüfe den vorherigen Absatz
            If Not para.Previous Is Nothing Then
                Set prevPara = para.Previous
                If prevPara.Style <> "SuS_Mengentext" And prevPara.Style <> "SuS_Kastenheadline" And prevPara.Style <> "SuS_Kastentext" Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Absatz mit Format 'SuS_Bilddateiname' muss stets ein Absatz mit diesen Formaten vorangehen: SuS_Mengentext, SuS_Kastenheadline, SuS_Kastentext [" & first40Chars & "]"
                End If
            End If

            ' Überprüfe den nächsten Absatz
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                If nextPara.Style = "SuS_Kastentext" Then
                    isValidCondition = False
                    If Not para.Previous Is Nothing Then
                        Set prevPara = para.Previous
                        If prevPara.Style = "SuS_Kastenheadline" Then
                            headlineText = prevPara.Range.Text
                            headlineText = Replace(headlineText, Chr(13), "")
                            headlineText = Replace(headlineText, Chr(10), "")
                            headlineText = Trim(headlineText)
                            If InStr(1, headlineText, "Porträt", vbTextCompare) > 0 Then
                                isValidCondition = True
                            End If
                        End If
                    End If
                    If Not isValidCondition Then
                        first40Chars = First40Characters(para)
                        AddLogEntry "Nach Format 'SuS_Bilddateiname' darf 'SuS_Kastentext' nur folgen, wenn der vorherige Absatz eine 'SuS_Kastenheadline' mit Inhalt 'Porträt' ist: [" & first40Chars & "]"
                    End If
                ElseIf nextPara.Style <> "SuS_Bild/Tabellenunterschrift" And nextPara.Style <> "SuS_Autor_Kurzbiografie" Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat 'SuS_Bilddateiname' muss stets eines dieser Absatzformate folgen: SuS_Bild/Tabellenunterschrift,SuS_Autor_Kurzbiografie,SuS_Kastentext (mit Bedingung) [" & first40Chars & "]"
                End If
            End If
        End If

        ' Für check_style_before_odd (SuS_Kastenheadline, multiStyles = "SuS_Mengentext")
        If para.Style = "SuS_Kastenheadline" Then
            styleCount = styleCount + 1
            If styleCount Mod 2 <> 0 Then
                If Not para.Previous Is Nothing Then
                    Set previousPara = para.Previous
                    If Not previousPara Is Nothing Then
                        first40Chars = First40Characters(para)
                        correctFormat = False
                        If previousPara.Style = "SuS_Mengentext" Then
                            correctFormat = True
                        End If
                        If correctFormat = False Then
                            AddLogEntry "Ungeradzahligen Vorkommen von Absätzen mit Format " & "SuS_Kastenheadline" & " muss stets ein Absatz mit diesen Formaten vorangehen: " & "SuS_Mengentext" & " [" & first40Chars & "]"
                        End If
                    End If
                End If
            End If
        End If

        ' FindParagraphAfterMustBe für SuS_Subhead2 -> SuS_Autorname
        If para.Style = "SuS_Subhead2" Then
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                correctFormat = False
                If nextPara.Style = "SuS_Autorname" Then
                    correctFormat = True
                End If
                If correctFormat = False Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat " & "SuS_Subhead2" & " muss stets eines dieser Absatzformate folgen: " & "SuS_Autorname" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' Für SuS_Bild/Tabellenunterschrift -> multiStyles
        If para.Style = "SuS_Bild/Tabellenunterschrift" Then
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                correctFormat = False
                aStyleList = Split("SuS_Mengentext,SuS_Kastentext,SuS_Absatzheadline,SuS_Unter_Absatzheadline,SuS_Kasten_Absatzheadline", ",")
                For counter = LBound(aStyleList) To UBound(aStyleList)
                    If nextPara.Style = aStyleList(counter) Then
                        correctFormat = True
                        Exit For
                    End If
                Next
                If correctFormat = False Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat " & "SuS_Bild/Tabellenunterschrift" & " muss stets eines dieser Absatzformate folgen: " & "SuS_Mengentext,SuS_Kastentext,SuS_Absatzheadline,SuS_Unter_Absatzheadline,SuS_Kasten_Absatzheadline" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' Für SuS_Links_und_Literatur_Headline -> SuS_Links_und_Literatur_Text
        If para.Style = "SuS_Links_und_Literatur_Headline" Then
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                correctFormat = False
                If nextPara.Style = "SuS_Links_und_Literatur_Text" Then
                    correctFormat = True
                End If
                If correctFormat = False Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat " & "SuS_Links_und_Literatur_Headline" & " muss stets eines dieser Absatzformate folgen: " & "SuS_Links_und_Literatur_Text" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' Für SuS_Links_und_Literatur_Text -> SuS_Links_und_Literatur_Text
        If para.Style = "SuS_Links_und_Literatur_Text" Then
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                correctFormat = False
                If nextPara.Style = "SuS_Links_und_Literatur_Text" Then
                    correctFormat = True
                End If
                If correctFormat = False Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat " & "SuS_Links_und_Literatur_Text" & " muss stets eines dieser Absatzformate folgen: " & "SuS_Links_und_Literatur_Text" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' FindParagraphAfterMustNotBe für SuS_Kastenheadline -> SuS_Kastenheadline
        If para.Style = "SuS_Kastenheadline" Then
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
                correctFormat = False
                If nextPara.Style = "SuS_Kastenheadline" Then
                    correctFormat = True
                End If
                If correctFormat = True Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Auf Absatzformat " & "SuS_Kastenheadline" & " darf keines dieser Absatzformate folgen: " & "SuS_Kastenheadline" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' FindParagraphBeforeMustNotBe für SuS_Kastenheadline -> SuS_Kastenheadline
        If para.Style = "SuS_Kastenheadline" Then
            If Not para.Previous Is Nothing Then
                Set previousPara = para.Previous
                correctFormat = False
                If previousPara.Style = "SuS_Kastenheadline" Then
                    correctFormat = True
                End If
                If correctFormat = True Then
                    first40Chars = First40Characters(para)
                    AddLogEntry "Absatz mit Format " & "SuS_Kastenheadline" & " darf nie ein Absatz mit diesen Formaten vorangehen: " & "SuS_Kastenheadline" & " [" & first40Chars & "]"
                End If
            End If
        End If

        ' check_odd_kastenheadline_empty
        If para.Style = "SuS_Kastenheadline" Then
            If styleCount Mod 2 <> 0 Then
                paraText = para.Range.Text
                paraText = Trim(paraText)
                Do While Right(paraText, 1) = Chr(13) Or Right(paraText, 1) = Chr(10)
                    paraText = Left(paraText, Len(paraText) - 1)
                Loop
                If paraText = "" Then
                    If Not para.Next Is Nothing Then
                        first40Chars = First40Characters(para.Next)
                    Else
                        first40Chars = "No next paragraph"
                    End If
                    AddLogEntry "Ungerades Vorkommen von 'SuS_Kastenheadline' darf nicht leer sein: [" & first40Chars & "]. Ausname: Cheatsheets/Infografiken"
                End If
            End If
        End If

        ' check_listing_followed_by_quellcode
        If para.Style = "SuS_Kastenheadline" Then
            If styleCount Mod 2 <> 0 Then
                headlineText = Replace(para.Range.Text, Chr(13), "")
                headlineText = Replace(headlineText, Chr(10), "")
                headlineText = Trim(headlineText)
                If InStr(1, headlineText, "Listing", vbTextCompare) > 0 Then
                    Set p = para.Next
                    Do While Not p Is Nothing And p.Style <> "SuS_Kastenheadline"
                        If p.Style <> "SuS_Quellcode" Then
                            offendingText = Replace(p.Range.Text, Chr(13), "")
                            offendingText = Replace(offendingText, Chr(10), "")
                            offendingText = Trim(offendingText)
                            AddLogEntry "Ungültiges Absatzformat in Codelisting (" & headlineText & ") gefunden: 'SuS_Quellcode' erwartet aber '" & p.Style & "' gefunden: " & offendingText
                        End If
                        Set p = p.Next
                    Loop
                End If
            End If
        End If
    Next para

    ' Log invalid styles
    If invalidStyles.Count > 0 Then
        For Each styleName In invalidStyles
            wrtstring = "Ungültiges Absatz- oder Zeichenformat gefunden: " & styleName
            AddLogEntry wrtstring
        Next styleName
    End If

    ' Perform style count checks
    Dim i As Long
    Dim found As Boolean
    Dim cnt As Long

    ' CHECK FOR SuS_Headline (only one)
    found = False
    cnt = 0
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Headline" Then
            cnt = styleCounts(i)
            found = True
            Exit For
        End If
    Next
    If found Then
        If cnt <> 1 Then
            AddLogEntry "Absatzformat SuS_Headline darf nur 1 mal vorkommen. Wird aber (" & cnt & ") mal verwendet."
        End If
    Else
        AddLogEntry "Absatzformat SuS_Headline darf nur 1 mal vorkommen. Wird aber (0) mal verwendet."
    End If

    ' CHECK FOR SuS_Subhead1 (only one)
    found = False
    cnt = 0
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Subhead1" Then
            cnt = styleCounts(i)
            found = True
            Exit For
        End If
    Next
    If found Then
        If cnt <> 1 Then
            AddLogEntry "Absatzformat SuS_Subhead1 darf nur 1 mal vorkommen. Wird aber (" & cnt & ") mal verwendet."
        End If
    Else
        AddLogEntry "Absatzformat SuS_Subhead1 darf nur 1 mal vorkommen. Wird aber (0) mal verwendet."
    End If

    ' CHECK FOR SuS_Autorname (only one)
    found = False
    cnt = 0
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Autorname" Then
            cnt = styleCounts(i)
            found = True
            Exit For
        End If
    Next
    If found Then
        If cnt <> 1 Then
            AddLogEntry "Absatzformat SuS_Autorname darf nur 1 mal vorkommen. Wird aber (" & cnt & ") mal verwendet."
        End If
    Else
        AddLogEntry "Absatzformat SuS_Autorname darf nur 1 mal vorkommen. Wird aber (0) mal verwendet."
    End If

    ' CHECK FOR SuS_Subhead2 (less than two)
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Subhead2" Then
            If styleCounts(i) >= 2 Then
                AddLogEntry "Absatzformat SuS_Subhead2 darf nur 1 mal (oder gar nicht) vorkommen. Wird aber (" & styleCounts(i) & ") mal verwendet."
            End If
            Exit For
        End If
    Next

    ' CHECK FOR SuS_Links_und_Literatur_Headline (less than two)
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Links_und_Literatur_Headline" Then
            If styleCounts(i) >= 2 Then
                AddLogEntry "Absatzformat SuS_Links_und_Literatur_Headline darf nur 1 mal (oder gar nicht) vorkommen. Wird aber (" & styleCounts(i) & ") mal verwendet."
            End If
            Exit For
        End If
    Next

    ' CHECK FOR SuS_Kastenheadline (modulo even)
    For i = 0 To styleCountIndex - 1
        If styleNames(i) = "SuS_Kastenheadline" Then
            If styleCounts(i) Mod 2 <> 0 Then
                AddLogEntry "Absatzformat SuS_Kastenheadline wurde nicht korrekt geschlossen. Bitte alle (" & styleCounts(i) & ") Vorkommen prüfen."
            End If
            Exit For
        End If
    Next

    ' 2) Zeichenformate: für jeden Zeichenstil im Dokument einmal Find.Execute
    For Each st In ActiveDocument.Styles
        If st.Type = wdStyleTypeCharacter Then
            Set searchRange = ActiveDocument.Content
            With searchRange.Find
                .ClearFormatting
                .Style = st
                .Text = ""
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
            End With
            ' Nur ein einziges Execute – wir wollen nur wissen OB der Stil vorkommt
            If searchRange.Find.Execute Then
                styleName = st.NameLocal
                If Left(styleName, 4) <> "SuS_" Then
                    On Error Resume Next
                    invalidStyles.Add styleName, styleName
                    On Error GoTo 0
                End If
            End If
        End If
    Next st

    If DebugMode Then MsgBox "Ende: ConsolidatedParagraphChecks"
End Sub