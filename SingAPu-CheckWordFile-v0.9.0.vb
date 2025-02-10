' version 0.9.0p

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
' Dim logFile As Object
' Dim logFilePath As String
Dim logFileName As String
Dim NameContainsSpecialChars As Boolean
Dim char As String
Dim logFilePath As String
Dim logFile As Integer


Sub SingAPu_CheckWordFile()





'----------------------------------------------------------
'----- CHECK FILE NAME -----
'----------------------------------------------------------

    ' MsgBox "Launching CHECK FILE NAME"
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

    invalidChars = "!@#$%^&*()+={}[]|\:;""'<>,.?/~`" ' Hier definierst du die Sonderzeichen, die du überprüfen möchtest.
    umlautChars = "äöüÄÖÜß" ' Umlaute und Sonderzeichen, die überprüft werden sollen.
    emptySpaceChar = " " ' Leerzeichen wird separat aufgeführt, um eine aussagekräftige Fehlermeldung zu generieren
    
    foundInvalid = False
    invalidList = "" ' Leere Liste für ungültige Zeichen
    
    ' Überprüfen des Dateinamens auf unerlaubte Sonderzeichen
    For i = 1 To Len(baseFileName)
        currentChar = Mid(baseFileName, i, 1)
        
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
    Next i
    
    ' Falls keine ungültigen Zeichen gefunden wurden
    If foundInvalid Then
        MsgBox "Die folgenden Sonderzeichen wurden im Dateinamen gefunden und müssen zunächst ersetzt werden: " & vbCrLf & invalidList, vbExclamation
        Exit Sub
    End If
    ' MsgBox "CHECK FILE NAME done"





'----------------------------------------------------------
'----- DELETE LOG FILE, IF EXISTS -----
'----------------------------------------------------------
' MsgBox "Launching DELETE LOG FILE"

' Set name of log file
logFileName = "log" & ".txt"
 
' Set the path for the log file
logFilePath = ActiveDocument.Path & "/" & logFileName
' MsgBox "logFilePath: " & logFilePath

If Dir(logFilePath) <> "" Then
    ' MsgBox "Log-Datei besteht bereits. Wird gelöscht."
    Kill logFilePath
End If

' MsgBox "DELETE LOG FILE done"





'----------------------------------------------------------
'----- CHECK IF FIRST PARAGRAPH HAS A PIPE IN IT, IF SO CHECK IF PARAGRAPH IS FORMAT X, IF NOT SET THE CORRECT FORMAT -----
'----------------------------------------------------------

Dim firstParagraph As Paragraph
Dim appliedStyle As String
Dim paraText As String

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





'----------------------------------------------------------
'----- SET HEADER FORMATS -----
'----------------------------------------------------------

' MsgBox "Launching SET HEADER FORMATS"
' SET FORMAT OF FIRST PARAGRAPH
' ActiveDocument.Paragraphs.First.Range.Select
' search_firstPara
'Format will be set in Sub if check is TRUE

' SET FORMAT OF SECOND PARAGRAPH
ActiveDocument.Paragraphs(2).Style = "SuS_Headline"

' SET FORMAT OF THIRD PARAGRAPH
ActiveDocument.Paragraphs(3).Style = "SuS_Subhead1"

' MsgBox "SET HEADER FORMATS done"





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
multiStyles = "SuS_Mengentext,SuS_Kastentext,SuS_Absatzheadline,SuS_Unter_Absatzheadline"
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

Dim para As Paragraph
Dim regex As Object
Dim specialCharPattern As String
Dim foundSpecialChar As Boolean

' Pattern für Sonderzeichen: alles außer Buchstaben, Zahlen und Unterstrich
specialCharPattern = "[^a-z0-9_]"

' Erstellen des regulären Ausdrucks-Objekts
Set regex = CreateObject("VBScript.RegExp")
regex.IgnoreCase = False  ' Groß-/Kleinschreibung beachten (da nur Kleinbuchstaben erlaubt)
regex.Global = True
regex.Pattern = specialCharPattern

' Durchlaufen aller Absätze im Dokument
For Each para In ActiveDocument.Paragraphs
    paraText = para.Range.Text

    ' Überprüfen, ob der Absatz das Format "SuS_Bilddateiname" enthält
    If InStr(paraText, "SuS_Bilddateiname") = 1 Then
        ' Überprüfen, ob unerlaubte Zeichen vorhanden sind
        If regex.Test(paraText) Then
            foundSpecialChar = True
            MsgBox "Fehler: Sonderzeichen im Absatz gefunden: " & vbCrLf & paraText, vbCritical
            Exit Sub
        End If
    End If
Next para





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
count_style_modulo





'----------------------------------------------------------
'----- END OF SCRIPT MESSAGE -----
'----------------------------------------------------------

MsgBox "Ich habe fertig."

End Sub





Sub count_style_onlyone()
' MsgBox "Sub count_style_onlyone() gestartet für Absatzformat: " & NameOfFormat
Dim l As Integer
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
End Sub


Sub count_style_lessthantwo()
' MsgBox "Sub count_style_lessthantwo() gestartet für Absatzformat: " & NameOfFormat
Dim l As Integer
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
End Sub


Sub count_style_modulo()
' MsgBox "Sub count_style_modulo() gestartet für Absatzformat: " & NameOfFormat
Dim l As Integer
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
    'MsgBox "Modulo = 0 – alle Kästen werden auch geschlossen."
Else
    WriteLogFile "Absatzformat " & NameOfFormat & " wurde nicht korrekt geschlossen. Bitte alle (" & l & ") Vorkommen prüfen."
End If
reset_search
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

    ' MsgBox "Funktion 'FindParagraphBeforeMustBe' gestartet für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    Dim ParaIndex As Long
    Dim ParaBefore As Long
    FindParagraphBeforeMustBe = ParaIndex

    For ParaIndex = 1 To SearchRange.Paragraphs.Count
        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then
            'jump 1 paragraph back and check if it has certain format
            ParaBefore = ParaIndex - 1
            char = ActiveDocument.Paragraphs(ParaIndex).Range.Sentences(1).Text
            ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
            correctFormat = False
            aStyleList = Split(multiStyles, ",")
            ' MsgBox "multiStyles: " & multiStyles
            For counter = LBound(aStyleList) To UBound(aStyleList)
                ' MsgBox "counter: " & counter
                NameOfFormatBefore = aStyleList(counter)
                ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                If ActiveDocument.Paragraphs(ParaBefore).Range.Style = NameOfFormatBefore Then
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                    correctFormat = True
                Else
                ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                End If
                ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Absatz mit Format " & ParaStyle & " muss stets ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & char & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
        End If
    Next
End Function





Public Function FindParagraphBeforeMustNotBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    ' MsgBox "Funktion 'FindParagraphBeforeMustNotBe' gestartet für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    Dim ParaIndex As Long
    Dim ParaBefore As Long
    FindParagraphBeforeMustNotBe = ParaIndex

    For ParaIndex = 1 To SearchRange.Paragraphs.Count
        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then
            'jump 1 paragraph back and check if it has certain format
            ParaBefore = ParaIndex - 1
            char = ActiveDocument.Paragraphs(ParaIndex).Range.Sentences(1).Text
            ' set Variable to FALSE – only gets TRUE if correct format is being used (see IF-statement)
            correctFormat = False
            aStyleList = Split(multiStyles, ",")
            ' MsgBox "multiStyles: " & multiStyles
            For counter = LBound(aStyleList) To UBound(aStyleList)
                ' MsgBox "counter: " & counter
                NameOfFormatBefore = aStyleList(counter)
                ' MsgBox "NameOfFormatBefore: " & NameOfFormatBefore
                If ActiveDocument.Paragraphs(ParaBefore).Range.Style <> NameOfFormatBefore Then
                    ' MsgBox "Absatz mit Format " & NameOfFormat & " geht korrekterweise Absatz mit Format " & NameOfFormatBefore & " voran."
                    correctFormat = True
                Else
                ' MsgBox "Absatz mit Format " & NameOfFormat & " geht nicht Absatz mit Format " & NameOfFormatBefore & " voran."
                End If
                ' MsgBox "Durchlauf für " & NameOfFormatBefore & " beendet. correctFormat: " & correctFormat
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Absatz mit Format " & ParaStyle & " darf nie ein Absatz mit diesen Formaten vorangehen: " & multiStyles & " [" & char & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
        End If
    Next
End Function



Public Function FindParagraphAfterMustBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    MsgBox "Funktion 'FindParagraphAfterMustBe' gestartet für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    Dim ParaIndex As Long
    Dim ParaAfter As Long
    FindParagraphAfterMustBe = ParaIndex
    ' ParaAfter = ParaIndex + 1
    For ParaIndex = 1 To SearchRange.Paragraphs.Count - 1
        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then
            'jump 1 paragraph ahead and check if it has certain format
            ParaAfter = ParaIndex + 1
            ' MsgBox "Format " & ParaStyle & " gefunden [" & ParaIndex & " + " & ParaAfter & "]."
            char = ActiveDocument.Paragraphs(ParaIndex).Range.Sentences(1).Text
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
                If ActiveDocument.Paragraphs(ParaAfter).Range.Style = NameOfFormatAfter Then
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise Format " & NameOfFormatAfter
                    correctFormat = True
                Else
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt nicht Format " & NameOfFormatAfter
                End If
            MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Auf Absatzformat " & ParaStyle & " muss stets eines dieser Absatzformate folgen: " & multiStyles & " [" & char & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
        End If
    Next
End Function




Public Function FindParagraphAfterMustNotBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    MsgBox "Funktion 'FindParagraphAfterMustNotBe' gestartet für Absatzformat: " & NameOfFormat & " (" & ParaStyle & ")"

    Dim ParaIndex As Long
    Dim ParaAfter As Long
    FindParagraphAfterMustNotBe = ParaIndex
    ' ParaAfter = ParaIndex + 1
    For ParaIndex = 1 To SearchRange.Paragraphs.Count - 1
        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then
            'jump 1 paragraph ahaed and check if it has certain format
            ParaAfter = ParaIndex + 1
            ' MsgBox "Format " & ParaStyle & " gefunden [" & ParaIndex & " + " & ParaAfter & "]."
            char = ActiveDocument.Paragraphs(ParaIndex).Range.Sentences(1).Text
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
                If ActiveDocument.Paragraphs(ParaAfter).Range.Style <> NameOfFormatAfter Then
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise keines der ausgeschlossenen Formate " & NameOfFormatAfter
                    correctFormat = True
                Else
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt fälschlicherweise eines der ausgeschlossenen Formate " & NameOfFormatAfter
                End If
            MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
            Next
            ' check if variable is FALSE and if so, write to logfile
            If correctFormat = False Then
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Auf Absatzformat " & ParaStyle & " darf keines dieser Absatzformate folgen: " & multiStyles & " [" & char & "]"
            End If
            ' MsgBox "correctFormat: " & correctFormat
        End If
    Next
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
