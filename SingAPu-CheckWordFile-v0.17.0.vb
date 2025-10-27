' version 0.17.0

'----------------------------------------------------------
'----- SET GLOBAL VARIABLES -----
'----------------------------------------------------------
Option Explicit
Public NameOfFormat As String
Public NameOfFormatAfter As String
Public NameOfFormatBefore As String
Dim multiStyles As String, I As Integer
Dim aStyleList As Variant
Dim counter As Long, s As String
Dim correctFormat As Boolean
Dim logFileName As String
Dim NameContainsSpecialChars As Boolean
Dim char As String
Dim logFilePath As String
Dim logFile As Integer
Dim para As Paragraph
Dim paraText As String
Dim first40Chars As String
Dim previousPara As Paragraph
Dim nextPara As Paragraph
Dim styleCount As Integer
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
    Dim logFileNumber As Integer
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

'----------------------------------------------------------
'----- CHECK FILE NAME -----
'----------------------------------------------------------

    ' MsgBox "Launching: CHECK FILE NAME"
    Dim fileName As String
    Dim baseFileName As String
    Dim invalidChars As String
    Dim i As Integer
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
    ' MsgBox "Done: CHECK FILE NAME"





'----------------------------------------------------------
'----- DELETE LOG FILE, IF EXISTS -----
'----------------------------------------------------------
' MsgBox "Launching: DELETE LOG FILE"

' Set name of log file
logFileName = "log" & ".txt"
 
' Set the path for the log file
logFilePath = ActiveDocument.Path & "/" & logFileName
' MsgBox "logFilePath: " & logFilePath

If Dir(logFilePath) <> "" Then
    ' MsgBox "Log-Datei besteht bereits. Wird gelöscht."
    Kill logFilePath
End If

' MsgBox "Done: DELETE LOG FILE"





'----------------------------------------------------------
'----- CHECK IF FIRST PARAGRAPH HAS A PIPE IN IT, IF SO CHECK IF PARAGRAPH IS FORMAT X, IF NOT SET THE CORRECT FORMAT -----
'----------------------------------------------------------

' MsgBox "Launching: CHECK FOR PIPE"

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

' MsgBox "Done: CHECK FOR PIPE"





'----------------------------------------------------------
'----- SET HEADER FORMATS -----
'----------------------------------------------------------

' MsgBox "Launching: SET HEADER FORMATS"

' SET FORMAT OF FIRST PARAGRAPH
' ActiveDocument.Paragraphs.First.Range.Select
' search_firstPara
'Format will be set in Sub if check is TRUE

' SET FORMAT OF SECOND PARAGRAPH
ActiveDocument.Paragraphs(2).Style = "SuS_Headline"

' SET FORMAT OF THIRD PARAGRAPH
ActiveDocument.Paragraphs(3).Style = "SuS_Subhead1"

' MsgBox "Done: SET HEADER FORMATS"





'----------------------------------------------------------
'----- CHECK NUMBER OF FORMAT INSTANCES  -----
'----------------------------------------------------------

' CHECK FOR SuS_Headline
NameOfFormat = "SuS_Headline"
count_style_onlyone

' CHECK FOR SuS_Subhead1
NameOfFormat = "SuS_Subhead1"
count_style_onlyone

' CHECK FOR SuS_Autorname
NameOfFormat = "SuS_Autorname"
count_style_onlyone

' CHECK FOR SuS_Autorname
NameOfFormat = "SuS_Subhead2"
count_style_lessthantwo

' CHECK FOR SuS_Autorname
NameOfFormat = "SuS_Links_und_Literatur_Headline"
count_style_lessthantwo





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER NEXT PARAGRAPH IS FORMAT Y -----
'----------------------------------------------------------
    
Dim IsFound As Boolean

NameOfFormat = "SuS_Subhead2"
multiStyles = "SuS_Autorname"
IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Bild/Tabellenunterschrift"
multiStyles = "SuS_Mengentext,SuS_Kastentext,SuS_Absatzheadline,SuS_Unter_Absatzheadline,SuS_Kasten_Absatzheadline"
IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Bilddateiname"
multiStyles = "SuS_Bild/Tabellenunterschrift,SuS_Autor_Kurzbiografie"
IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Links_und_Literatur_Headline"
multiStyles = "SuS_Links_und_Literatur_Text"
IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Links_und_Literatur_Text"
multiStyles = "SuS_Links_und_Literatur_Text"
IsFound = FindParagraphAfterMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER NEXT PARAGRAPH IS NOT FORMAT Y -----
'----------------------------------------------------------

NameOfFormat = "SuS_Kastenheadline"
multiStyles = "SuS_Kastenheadline"
IsFound = FindParagraphAfterMustNotBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER PREVIOUS PARAGRAPH IS FORMAT Y -----
'----------------------------------------------------------

NameOfFormat = "SuS_Bilddateiname"
multiStyles = "SuS_Mengentext,SuS_Kastentext"
IsFound = FindParagraphBeforeMustBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER PREVIOUS PARAGRAPH IS NOT FORMAT Y -----
'----------------------------------------------------------

NameOfFormat = "SuS_Kastenheadline"
multiStyles = "SuS_Kastenheadline"
IsFound = FindParagraphBeforeMustNotBe(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)





'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER IT CONTAINS SPECIAL CHARACTERS -----
'----------------------------------------------------------

' MsgBox "Start: 'find special chars in image reference'"

Dim Absatz As Paragraph
Dim Formatvorlage As String
Dim baseAbsatzText As String
Dim j As Integer
Dim dotPos As Integer
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

' MOD Dividiert zwei Zahlen und gibt nur den Rest zurück.
'Result = number1 Mod number2
'number1 wäre die Anzahl der Vorkommen von Format X
'number2 wäre 2
' Wenn Rest gleich 0, dann ist es ein ganzzahliges Vielfaches von 2 – wird also geöffnet UND geschlossen

' CHECK FOR SuS_Kastenheadline
NameOfFormat = "SuS_Kastenheadline"
count_style_modulo_even





'----------------------------------------------------------
'----- CHECK IF PREVIOUS PARAGRAPH OF ODD OCCURENCES OF FORMAT X IS FORMAT Y -----
'----------------------------------------------------------

' CHECK FOR SuS_Kastenheadline
NameOfFormat = "SuS_Kastenheadline"
multiStyles = "SuS_Mengentext"
check_style_before_odd





'----------------------------------------------------------
'----- CHECK IF ODD OCCURENCES OF FORMAT X IS NOT EMPTY ---
'----------------------------------------------------------

NameOfFormat = "SuS_Kastenheadline"
check_odd_kastenheadline_empty

' NEW: Prüfe bei ungeraden SuS_Kastenheadline mit "Listing", dass alle folgenden Absätze bis zur nächsten SuS_Kastenheadline SuS_Quellcode sind
check_listing_followed_by_quellcode

'----------------------------------------------------------
'----- CHECK IF ONLY STYLES WITH STRING 'SuS_' ARE BEING USED ---
'----------------------------------------------------------
check_invalid_styles





'----------------------------------------------------------
'----- WRITE ALL LOG ENTRIES TO THE FILE ------------------
'----------------------------------------------------------

    WriteLogEntries





'----------------------------------------------------------
'----- END OF SCRIPT MESSAGE -----
'----------------------------------------------------------

MsgBox "Ich habe fertig."

End Sub





Sub count_style_onlyone()
    ' MsgBox "Start: Sub count_style_onlyone() für Absatzformat: " & NameOfFormat
    Dim l As Integer
    l = 0
    reset_search
    With ActiveDocument.Range.Find
    .Style = NameOfFormat 'Replace with the name of the style you are counting
    While .Execute
        l = l + 1
        If l > ActiveDocument.Range.Paragraphs.Count Then
            Stop
        End If
    Wend
    End With
    If l = 1 Then
        'MsgBox NameOfFormat & " passt"
    Else
        AddLogEntry "Absatzformat " & NameOfFormat & " darf nur 1 mal vorkommen. Wird aber " & "(" & l & ") mal verwendet."
        'MsgBox "Absatzformat " & NameOfFormat & " darf nur 1 mal vorkommen. Wird aber " & "(" & l & ") mal verwendet."
    End If
    reset_search
    ' MsgBox "Ende: Sub count_style_onlyone() für Absatzformat: " & NameOfFormat
End Sub


Sub count_style_lessthantwo()
    ' MsgBox "Start: Sub count_style_lessthantwo() für Absatzformat: " & NameOfFormat
    Dim l As Integer
    l = 0
    reset_search
    With ActiveDocument.Range.Find
    .Style = NameOfFormat 'Replace with the name of the style you are counting
    While .Execute
        l = l + 1
        If l > ActiveDocument.Range.Paragraphs.Count Then
            Stop
        End If
    Wend
    End With
    If l < 2 Then
        'MsgBox NameOfFormat & " passt"
    Else
        AddLogEntry "Absatzformat " & NameOfFormat & " darf nur 1 mal (oder gar nicht) vorkommen. Wird aber " & "(" & l & ") mal verwendet."
        'MsgBox "Absatzformat " & NameOfFormat & " darf nur 1 mal vorkommen. Wird aber " & "(" & l & ") mal verwendet."
    End If
    reset_search
    ' MsgBox "Ende: Sub count_style_lessthantwo() für Absatzformat: " & NameOfFormat
End Sub


Sub count_style_modulo_even()
    ' MsgBox "Start: Sub count_style_modulo_even() für Absatzformat: " & NameOfFormat
    Dim l As Integer
    l = 0
    reset_search
    With ActiveDocument.Range.Find
    .Style = NameOfFormat 'Replace with the name of the style you are counting
    While .Execute
        l = l + 1
        If l > ActiveDocument.Range.Paragraphs.Count Then
            Stop
        End If
    Wend
    End With

    Dim formatClosed As Integer
    Dim number1 As Integer
    Dim number2 As Integer

    number1 = l 'number of instances found
    number2 = 2 'integer multiple of 2 (opened + closed)

    formatClosed = number1 Mod number2

    If formatClosed = 0 Then
        'MsgBox "Modulo = 0 –> alle Kästen werden auch geschlossen."
    Else
        AddLogEntry "Absatzformat " & NameOfFormat & " wurde nicht korrekt geschlossen. Bitte alle (" & l & ") Vorkommen prüfen."
    End If
    reset_search
    ' MsgBox "Ende: Sub count_style_modulo_even() für Absatzformat: " & NameOfFormat
End Sub




Sub check_style_before_odd()

    ' MsgBox "Start: check_style_before_odd()"
    styleCount = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Style = NameOfFormat Then
            ' MsgBox "Absatz mit Format " & NameOfFormat & " wurde gefunden."
            styleCount = styleCount + 1
            ' MsgBox "styleCount: " & styleCount
            ' MsgBox "para.Range.Text: " & para.Range.Text

            ' Überprüfe nur ungerade Vorkommen
            If styleCount Mod 2 <> 0 Then
                ' Wenn es ungerade ist, überprüfe den vorherigen Absatz
                ' MsgBox "Ungerades Vorkommen von " & NameOfFormat & " wurde gefunden."
                If Not para.Previous Is Nothing Then
                    Set previousPara = para.Previous
                End If
                If Not previousPara Is Nothing Then ' Validate before accessing `previousPara.Style`
                    ' MsgBox "previousPara.Range.Text: " & previousPara.Range.Text
                    ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
                    first40Chars = First40Characters(para)
                    correctFormat = False
                    aStyleList = Split(multiStyles, ",")
                    ' MsgBox "multiStyles: " & multiStyles
                    For counter = LBound(aStyleList) To UBound(aStyleList)
                        ' MsgBox "counter: " & counter
                        NameOfFormatBefore = aStyleList(counter)
                        ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                        If previousPara.Style = NameOfFormatBefore Then
                            ' MsgBox "Ungerades Vorkommen von " & NameOfFormat & " gefunden, ohne 'SuS_Mengentext' davor!"
                            ' Rufe die Funktion auf, um die ersten 40 Zeichen des Absatzes zu holen
                            correctFormat = True
                        End If
                    Next
                    ' MsgBox "correctFormat: " & correctFormat & vbCrLf & "Wenn TRUE, enthält Dokument keinen Fehler."
                    ' check if variable is FALSE and if so, write to logfile
                    If correctFormat = False Then
                        AddLogEntry "Ungeradzahligen Vorkommen von Absätzen mit Format " & NameOfFormat & " muss stets ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
                    End If
                End If
            End If
        End If
    Next para
    ' MsgBox "Done: check_style_before_odd()"

End Sub




Sub check_odd_kastenheadline_empty()

    ' Initialize the counter for the style
    styleCount = 0

    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph has the style "SuS_Kastenheadline"
        If para.Style = NameOfFormat Then
            ' Increment the counter for this style
            styleCount = styleCount + 1

            ' Check if it's an odd-numbered instance
            If styleCount Mod 2 <> 0 Then
                ' Get the text of the paragraph
                paraText = para.Range.Text

                ' Remove leading/trailing spaces and line breaks
                paraText = Trim(paraText)
                Do While Right(paraText, 1) = Chr(13) Or Right(paraText, 1) = Chr(10)
                    paraText = Left(paraText, Len(paraText) - 1)
                Loop

                ' If the paragraph is empty, log a note
                If paraText = "" Then
                    ' Get the first 40 characters of the previous paragraph
                    If Not para.Next Is Nothing Then
                        first40Chars = First40Characters(para.Next)
                    Else
                        first40Chars = "No next paragraph"
                    End If

                    ' Write to the log file
                    AddLogEntry "Ungerades Vorkommen von 'SuS_Kastenheadline' darf nicht leer sein: [" & first40Chars & "]. Ausname: Cheatsheets/Infografiken"
                End If
            End If
        End If
    Next para
End Sub





Sub check_invalid_styles()
    Dim charRange As Variant
    Dim styleName As Variant
    Dim invalidStyles As Collection
    Dim style As Style
    Dim wrtstring As String

    ' Initialize a collection to store invalid styles
    Set invalidStyles = New Collection

    ' Check all paragraph styles in the document
    For Each para In ActiveDocument.Paragraphs
        styleName = para.Style.NameLocal
        If Left(styleName, 4) <> "SuS_" Then
            On Error Resume Next
            invalidStyles.Add styleName, styleName ' Avoid duplicates
            On Error GoTo 0
        End If
    Next para

    ' Check all character styles in the document
    For Each charRange In ActiveDocument.StoryRanges(wdMainTextStory).Characters
        styleName = charRange.Style.NameLocal
        If Left(styleName, 4) <> "SuS_" Then
            On Error Resume Next
            invalidStyles.Add styleName, styleName ' Avoid duplicates
            On Error GoTo 0
        End If
    Next charRange

    ' Log invalid styles to the log file
    If invalidStyles.Count > 0 Then
        For Each styleName In invalidStyles
            wrtstring = "Ungültiges Absatz- oder Zeichenformat gefunden: " & styleName
            AddLogEntry wrtstring
        Next styleName
    End If
End Sub





'----------------------------------------------------------
'----- SUB FOR RESETTING THE SEARCH -----
'----------------------------------------------------------
Sub reset_search()
With Selection.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = ""
   .Replacement.Text = ""
   .Forward = True
   .Wrap = wdFindContinue
   .Format = False
   .MatchCase = False
   .MatchWholeWord = False
   .MatchWildcards = False
   .MatchSoundsLike = False
   .MatchAllWordForms = False
   ' plus some more if needed
   .Execute
End With
End Sub





Public Function FindParagraphBeforeMustBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    ' MsgBox "Start: 'FindParagraphBeforeMustBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    For Each para In ActiveDocument.Paragraphs
        If para.Style = NameOfFormat Then
            'jump 1 paragraph back and check if it has certain format
            If Not para.Previous Is Nothing Then
                Set previousPara = para.Previous
            End If
            If Not previousPara Is Nothing Then ' Validate before accessing `previousPara.Style`
                ' Den Text des Absatzes holen
                paraText = para.Range.Text
                ' Die ersten 40 Zeichen extrahieren
                first40Chars = Left(paraText, 40)
                ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
                correctFormat = False
                aStyleList = Split(multiStyles, ",")
                ' MsgBox "multiStyles: " & multiStyles
                For counter = LBound(aStyleList) To UBound(aStyleList)
                    ' MsgBox "counter: " & counter
                    NameOfFormatBefore = aStyleList(counter)
                    ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                    If previousPara.Style = NameOfFormatBefore Then
                        ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                        correctFormat = True
                    Else
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
                Next
                ' check if variable is FALSE and if so, write to logfile
                If correctFormat = False Then
                    AddLogEntry "Absatz mit Format " & NameOfFormat & " muss stets ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
                End If
                ' MsgBox "correctFormat: " & correctFormat
            End If
        End If
    Next
    ' MsgBox "Ende: 'FindParagraphBeforeMustBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"
End Function





Public Function FindParagraphBeforeMustNotBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    ' MsgBox "Start: 'FindParagraphBeforeMustNotBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"
    For Each para In ActiveDocument.Paragraphs
        If para.Style = NameOfFormat Then
            ' Den Text des Absatzes holen
            paraText = para.Range.Text
            ' Die ersten 40 Zeichen extrahieren
            first40Chars = Left(paraText, 40)
            'jump 1 paragraph back and check if it has certain format
            If Not para.Previous Is Nothing Then
                Set previousPara = para.Previous
            End If
            If Not previousPara Is Nothing Then ' Validate before accessing `previousPara.Style`
                ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
                correctFormat = False
                aStyleList = Split(multiStyles, ",")
                ' MsgBox "multiStyles: " & multiStyles
                For counter = LBound(aStyleList) To UBound(aStyleList)
                    ' MsgBox "counter: " & counter
                    NameOfFormatBefore = aStyleList(counter)
                    ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                    If previousPara.Style = NameOfFormatBefore Then
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                    correctFormat = True
                    Else
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
                Next
                ' check if variable is FALSE and if so, write to logfile
                If correctFormat = True Then
                    AddLogEntry "Absatz mit Format " & NameOfFormat & " darf nie ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
                End If
                ' MsgBox "correctFormat: " & correctFormat
            End If
        End If
    Next
    ' MsgBox "Ende: 'FindParagraphBeforeMustNotBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"
End Function



Public Function FindParagraphAfterMustBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    ' MsgBox "Start: 'FindParagraphAfterMustBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    For Each para In ActiveDocument.Paragraphs
        If para.Style = NameOfFormat Then
            ' Den Text des Absatzes holen
            paraText = para.Range.Text
            ' Die ersten 40 Zeichen extrahieren
            first40Chars = Left(paraText, 40)
            'jump 1 paragraph ahead and check if it has certain format
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
            End If
            If Not nextPara Is Nothing Then ' Validate before accessing `nextPara.Style`
                ' MsgBox "Format " & ParaStyle & " gefunden [" & ParaIndex & " + " & ParaAfter & "]."
                ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
                correctFormat = False
                ' MsgBox "Variable correctFormat wird zunächst auf FALSE gesetzt: " & correctFormat
                aStyleList = Split(multiStyles, ",")
                ' MsgBox "multiStyles: " & multiStyles
                For counter = LBound(aStyleList) To UBound(aStyleList)
                    ' MsgBox "counter: " & counter
                    NameOfFormatAfter = aStyleList(counter)
                    ' MsgBox "NameOfFormatAfter: " & NameOfFormatAfter
                    If nextPara.Style = NameOfFormatAfter Then
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise Format " & NameOfFormatAfter
                        correctFormat = True
                    Else
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt nicht Format " & NameOfFormatAfter
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
                Next
                ' MsgBox "correctFormat: " & correctFormat & vbCrLf & "Wenn TRUE, enthält Dokument keinen Fehler."
                ' check if variable is FALSE and if so, write to logfile
                If correctFormat = False Then
                    AddLogEntry "Auf Absatzformat " & NameOfFormat & " muss stets eines dieser Absatzformate folgen: " & multiStyles & " [" & first40Chars & "]"
                End If
                ' MsgBox "correctFormat: " & correctFormat
            End If
        End If
    Next
    ' MsgBox "Ende: 'FindParagraphAfterMustBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"
End Function




Public Function FindParagraphAfterMustNotBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    ' MsgBox "Start: 'FindParagraphAfterMustNotBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    For Each para In ActiveDocument.Paragraphs
        If para.Style = NameOfFormat Then
            ' Den Text des Absatzes holen
            paraText = para.Range.Text
            ' Die ersten 40 Zeichen extrahieren
            first40Chars = Left(paraText, 40)
            'jump 1 paragraph ahaed and check if it has certain format
            If Not para.Next Is Nothing Then
                Set nextPara = para.Next
            End If
            If Not nextPara Is Nothing Then ' Validate before accessing `nextPara.Style`
                ' MsgBox "Format " & ParaStyle & " gefunden [" & ParaIndex & " + " & ParaAfter & "]."
                ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
                correctFormat = False
                ' MsgBox "Variable correctFormat wird zunächst auf FALSE gesetzt: " & correctFormat
                ' Test array
                aStyleList = Split(multiStyles, ",")
                ' MsgBox "multiStyles: " & multiStyles
                For counter = LBound(aStyleList) To UBound(aStyleList)
                    ' MsgBox "counter: " & counter
                    NameOfFormatAfter = aStyleList(counter)
                    ' MsgBox "NameOfFormatAfter: " & NameOfFormatAfter
                    If nextPara.Style = NameOfFormatAfter Then
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise keines der ausgeschlossenen Formate " & NameOfFormatAfter
                        correctFormat = True
                    Else
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt fälschlicherweise eines der ausgeschlossenen Formate " & NameOfFormatAfter
                    End If
                ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
                Next
                ' check if variable is FALSE and if so, write to logfile
                If correctFormat = True Then
                    AddLogEntry "Auf Absatzformat " & NameOfFormat & " darf keines dieser Absatzformate folgen: " & multiStyles & " [" & first40Chars & "]"
                End If
                ' MsgBox "correctFormat: " & correctFormat
            End If
        End If
    Next
    ' MsgBox "Ende: 'FindParagraphAfterMustNotBe' für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"
End Function





Function First40Characters(para As Paragraph) As String
    ' MsgBox "Start: Funktion 'First40Characters'"
    
    ' Den Text des Absatzes holen
    paraText = para.Range.Text
    ' MsgBox "paraText: " & paraText
    
    ' Die ersten 40 Zeichen extrahieren
    first40Chars = Left(paraText, 40)
    ' MsgBox "first40Chars: " & first40Chars

    ' Entferne führende Leerzeichen und Zeilenumbrüche
    first40Chars = RTrim(first40Chars)

    ' Entferne Zeilenumbrüche und Leerzeichen am Ende des Textes
    Do While Right(first40Chars, 1) = Chr(13) Or Right(first40Chars, 1) = Chr(10)
        first40Chars = Left(first40Chars, Len(first40Chars) - 1)
    Loop
    
    ' Gib die ersten 40 Zeichen zurück
    First40Characters = first40Chars
End Function


Sub check_listing_followed_by_quellcode()
    Dim headlineCount As Long
    Dim hPara As Paragraph
    Dim p As Paragraph
    Dim headlineText As String

    headlineCount = 0

    For Each hPara In ActiveDocument.Paragraphs
        If hPara.Style = "SuS_Kastenheadline" Then
            headlineCount = headlineCount + 1

            ' Only check odd-numbered instances
            If headlineCount Mod 2 <> 0 Then
                ' Normalize headline text and check for "Listing"
                headlineText = Replace(hPara.Range.Text, Chr(13), "")
                headlineText = Replace(headlineText, Chr(10), "")
                headlineText = Trim(headlineText)

                If InStr(1, headlineText, "Listing", vbTextCompare) > 0 Then
                    ' Walk following paragraphs until the next SuS_Kastenheadline (closing)
                    Set p = hPara.Next
                    Do While Not p Is Nothing And p.Style <> "SuS_Kastenheadline"
                        If p.Style <> "SuS_Quellcode" Then
                            ' Prepare cleaned texts
                            Dim headlineTextClean As String
                            Dim offendingText As String
                            headlineTextClean = Replace(headlineText, Chr(13), "")
                            headlineTextClean = Replace(headlineTextClean, Chr(10), "")
                            headlineTextClean = Trim(headlineTextClean)
                            offendingText = Replace(p.Range.Text, Chr(13), "")
                            offendingText = Replace(offendingText, Chr(10), "")
                            offendingText = Trim(offendingText)
                            ' Log detailed message
                            AddLogEntry "Ungültiges Absatzformat in Codelisting (" & headlineTextClean & ") gefunden: 'SuS_Quellcode' erwartet aber '" & p.Style & "' gefunden: " & offendingText
                        End If
                        Set p = p.Next
                    Loop
                End If
            End If
        End If
    Next hPara
End Sub