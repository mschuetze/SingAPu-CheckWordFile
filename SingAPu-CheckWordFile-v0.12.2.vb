' version 0.12.2

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





Sub SingAPu_CheckWordFile()

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
    
    fileName = ActiveDocument.Name
    ' Entfernen der Dateiendung (alles nach dem letzten Punkt)
    baseFileName = Left(fileName, InStrRev(fileName, ".") - 1)
    ' MsgBox "baseFileName: " & baseFileName

    invalidChars = "!@#$%^&*()+={}[]|\:;""'<>,.?/~`" ' Hier definierst du die Sonderzeichen, die du überprüfen möchtest.
    umlautChars = "äöüÄÖÜß" ' Umlaute und Sonderzeichen, die überprüft werden sollen.
    emptySpaceChar = " " ' Leerzeichen wird separat aufgeführt, um eine aussagekräftige Fehlermeldung zu generieren
    
    foundInvalid = False
    invalidList = "" ' Leere Liste für ungültige Zeichen
    
    ' Überprüfen des Dateinamens auf unerlaubte Sonderzeichen
    For i = 1 To Len(baseFileName)
        currentChar = Mid(baseFileName, i, 1)
        ' MsgBox "currentChar: " & currentChar
        
        ' Überprüfen auf Sonderzeichen
        If InStr(invalidChars, currentChar) > 0 Then
            MsgBox "Der Dateiname (" & baseFileName & ") enthält ein Sonderzeichen"
            If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: " & currentChar, vbExclamation
                invalidList = invalidList & currentChar & " " ' Füge das ungültige Zeichen der Liste hinzu
            End If
            foundInvalid = True
        End If

        ' Überprüfen auf Umlaute
        If InStr(umlautChars, currentChar) > 0 Then
            MsgBox "Der Dateiname (" & baseFileName & ") enthält einen Umlaut"
            If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: " & currentChar, vbExclamation
                invalidList = invalidList & currentChar & " " ' Füge das ungültige Zeichen der Liste hinzu
            End If
            foundInvalid = True
        End If

        ' Überprüfen auf Leerzeichen
        If InStr(emptySpaceChar, currentChar) > 0 Then
            MsgBox "Der Dateiname (" & baseFileName & ") enthält ein Leerzeichen"
            If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                ' MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: Leerzeichen", vbExclamation
                invalidList = invalidList & "Leerzeichen " ' Füge das ungültige Zeichen der Liste hinzu
            End If
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

' Die Formatvorlage, die du suchen möchtest
Formatvorlage = "SuS_Bilddateiname"

' Definiere die gängigen Bild-Dateiendungen
Dateiendung = Array(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg")


' Gehe alle Absätze im Dokument durch
For Each Absatz In ActiveDocument.Paragraphs
    ' Wenn der Absatz die angegebene Formatvorlage hat
    If Absatz.Style = Formatvorlage Then
        foundInvalid = False
        EndungGefunden = False ' Standardmäßig auf ungültig setzen
        invalidList = "" ' Leere Liste für ungültige Zeichen
        ' Hole den Text des Absatzes
        paraText = Absatz.Range.Text
        ' MsgBox "paraText: " & paraText
        ' Schleife über alle gängigen Bilddateiendungen
        For i = LBound(Dateiendung) To UBound(Dateiendung)
            If InStr(1, paraText, Dateiendung(i), vbTextCompare) > 0 Then
                EndungGefunden = True
                Exit For
            End If
        Next i

        If Not EndungGefunden Then
            WriteLogFile "Dem Bildverweis '" & paraText & "' fehlt eine Dateiendung (.tif, .jpg, usw)"
        Else
            ' Entfernen der Dateiendung (alles nach dem letzten Punkt)
            baseAbsatzText = Left(paraText, InStrRev(paraText, ".") - 1)
            ' MsgBox "baseAbsatzText: " & baseAbsatzText

            ' Überprüfe jeden Charakter im Absatz auf Sonderzeichen
            For j = 1 To Len(baseAbsatzText)
                currentChar = Mid(baseAbsatzText, j, 1)

                ' Überprüfen auf Sonderzeichen
                If InStr(invalidChars, currentChar) > 0 Then
                    If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                        ' MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: " & currentChar, vbExclamation
                        invalidList = invalidList & currentChar & " " ' Füge das ungültige Zeichen der Liste hinzu
                    End If
                    foundInvalid = True
                End If

                ' Überprüfen auf Umlaute
                If InStr(umlautChars, currentChar) > 0 Then
                    If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                        ' MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: " & currentChar, vbExclamation
                        invalidList = invalidList & currentChar & " " ' Füge das ungültige Zeichen der Liste hinzu
                    End If
                    foundInvalid = True
                End If

                ' Überprüfen auf Leerzeichen
                If InStr(emptySpaceChar, currentChar) > 0 Then
                    If InStr(invalidList, currentChar) = 0 Then ' Verhindern von Duplikaten in der Liste
                        ' MsgBox "Der Dateiname enthält ein ungültiges Sonderzeichen: Leerzeichen", vbExclamation
                        invalidList = invalidList & "Leerzeichen " ' Füge das ungültige Zeichen der Liste hinzu
                    End If
                    foundInvalid = True
                End If

            Next j

            ' Wenn Sonderzeichen gefunden wurden, in log.txt vermerken
            If foundInvalid Then
                WriteLogFile "Bildverweis " & paraText & "enthält folgende Sonderzeichen: " & invalidList
            End If
        End If
    End If
Next Absatz

' MsgBox "Ende: 'find special chars in image reference'"





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

' CHECK FOR SuS_Kastenheadline
NameOfFormat = "SuS_Kastenheadline"
check_odd_kastenheadline_empty




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
        WriteLogFile "Absatzformat " & NameOfFormat & " darf nur 1 mal vorkommen. Wird aber " & "(" & l & ") mal verwendet."
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
        WriteLogFile "Absatzformat " & NameOfFormat & " darf nur 1 mal (oder gar nicht) vorkommen. Wird aber " & "(" & l & ") mal verwendet."
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
        WriteLogFile "Absatzformat " & NameOfFormat & " wurde nicht korrekt geschlossen. Bitte alle (" & l & ") Vorkommen prüfen."
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
                Set previousPara = para.Previous
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
                    If Not previousPara Is Nothing Then
                        If previousPara.Style = NameOfFormatBefore Then
                            ' MsgBox "Ungerades Vorkommen von " & NameOfFormat & " gefunden, ohne 'SuS_Mengentext' davor!"
                            ' Rufe die Funktion auf, um die ersten 40 Zeichen des Absatzes zu holen
                            correctFormat = True
                        End If
                    End If
                Next
                ' MsgBox "correctFormat: " & correctFormat & vbCrLf & "Wenn TRUE, enthält Dokument keinen Fehler."
                ' check if variable is FALSE and if so, write to logfile
                If correctFormat = False Then
                    WriteLogFile "Ungeradzahligen Vorkommen von Absätzen mit Format " & NameOfFormat & " muss stets ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
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
                    WriteLogFile "Ungerades Vorkommen von 'SuS_Kastenheadline' darf nicht leer sein: [" & first40Chars & "]."
                End If
            End If
        End If
    Next para
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
            Set previousPara = para.Previous
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
                If Not previousPara Is Nothing Then
                    If previousPara.Style = NameOfFormatBefore Then
                        ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                        correctFormat = True
                    Else
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
                End If
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Absatz mit Format " & NameOfFormat & " muss stets ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
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
            Set previousPara = para.Previous
            ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
            correctFormat = False
            aStyleList = Split(multiStyles, ",")
            ' MsgBox "multiStyles: " & multiStyles
            For counter = LBound(aStyleList) To UBound(aStyleList)
                ' MsgBox "counter: " & counter
                NameOfFormatBefore = aStyleList(counter)
                ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                If Not previousPara Is Nothing Then
                    If previousPara.Style = NameOfFormatBefore Then
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                    correctFormat = True
                    Else
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
                End If
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = True Then
                WriteLogFile "Absatz mit Format " & NameOfFormat & " darf nie ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & first40Chars & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
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
            Set nextPara = para.Next
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
                If Not nextPara Is Nothing Then
                    If nextPara.Style = NameOfFormatAfter Then
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise Format " & NameOfFormatAfter
                        correctFormat = True
                    Else
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt nicht Format " & NameOfFormatAfter
                    End If
                    ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
                Else
                    correctFormat = True
                End If
            Next
            ' MsgBox "correctFormat: " & correctFormat & vbCrLf & "Wenn TRUE, enthält Dokument keinen Fehler."
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Auf Absatzformat " & NameOfFormat & " muss stets eines dieser Absatzformate folgen: " & multiStyles & " [" & first40Chars & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
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
            Set nextPara = para.Next
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
                If Not nextPara Is Nothing Then
                    If nextPara.Style = NameOfFormatAfter Then
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise keines der ausgeschlossenen Formate " & NameOfFormatAfter
                        correctFormat = True
                    Else
                        ' MsgBox "Auf Format " & NameOfFormat & " folgt fälschlicherweise eines der ausgeschlossenen Formate " & NameOfFormatAfter
                    End If
                Else
                    correctFormat = False
                End If
            ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = True Then
                WriteLogFile "Auf Absatzformat " & NameOfFormat & " darf keines dieser Absatzformate folgen: " & multiStyles & " [" & first40Chars & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
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




Sub WriteLogFile(wrtstring As String)
    Dim logEntry As String
    Dim logFileNumber As Integer
 
    ' Create the log entry
    logEntry = Now & " - This is a log entry."
 
    ' Get a free file number
    logFileNumber = FreeFile
 
    ' Open the file for appending
    Open logFilePath For Append As #logFileNumber
 
    ' Write the log entry to the file
    Write #1, Now() & " : " & wrtstring & vbCrLf & "----" & vbCrLf
 
    ' Close the file
    Close #logFileNumber
End Sub
