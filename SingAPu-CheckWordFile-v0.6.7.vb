' version 0.6.7

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
Dim logFile As Object
Dim logFilePath As String
Dim logFileName As String
Dim NameContainsSpecialChars As Boolean
Dim char As String


Sub SingAPu_CheckWordFile()





'----------------------------------------------------------
'----- CHECK FILE NAME -----
'----------------------------------------------------------
' MsgBox "Launching CHECK FILE NAME"

Dim fileName As String
fileName = ActiveDocument.Name
NameContainsSpecialChars = False
'MsgBox fileName
BadChar (fileName)
If NameContainsSpecialChars = True Then
    MsgBox "Dateiname enthält Sonderzeichen."
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
'----- SET HEADER FORMATS -----
'----------------------------------------------------------

' MsgBox "Launching SET HEADER FORMATS"
' SET FORMAT OF FIRST PARAGRAPH
ActiveDocument.Paragraphs.First.Range.Select
search_firstPara
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



'----------------------------------------------------------
'----- CHECK FILE NAME -----
'----------------------------------------------------------
 
Function BadChar(strText As String) As Long
     '
     '****************************************************************************************
     '       Title       BadChar
     '       Target Application:  any
     '       Function    test for the presence of charcters that can not be used in
     '                   the name of an xlsheet, file, directory, etc
     '
     '           if no bad characters are found, BadChar = 0 on return
     '           if any bad character is found, BadChar = i where i is the index (in strText)
     '               where bad char was found
     '       Limitations:    passed string variable should not include any path seperator
     '                           characters
     '                       stops and exits when 1st bad char is found so # of bad chars
     '                           is not really known
     '       Passed Values:
     '           strText     [in, string]  text string to be examined
     '
     '****************************************************************************************
     '
     '
    Dim BadChars    As String
    Dim I           As Long
    Dim J           As Long
     
    ' Use Unicode for German Umlaute: https://stackoverflow.com/questions/22017723/regex-for-umlaut '
    BadChars = "ß:\/? *[](){}"
    'BadChars = ":\/?*[]"
    For I = 1 To Len(BadChars)
        J = InStr(strText, Mid(BadChars, I, 1))
        If J > 0 Then
            BadChar = J
            NameContainsSpecialChars = True
            Exit Function
        End If
    Next I
    BadChar = 0
     
End Function



Sub count_style_onlyone()
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


'----------------------------------------------------------
'----- SUB FOR SEARCHING THE FIRST PARAGRAPH -----
'----------------------------------------------------------

'This can be used to determine if the first paragraph really is the one with the pipe "|" in it
Sub search_firstPara()
    ' MsgBox "Launching search_firstPara()"
    Dim textToFind As String
    textToFind = " | "
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = textToFind
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
        Selection.Find.Execute
        If .Found = True Then
            ActiveDocument.Paragraphs.First.Style = "SuS_Mengentext"
        Else
            WriteLogFile "Fehler in Zeile 1: " & " Absatz muss eine Pipe ('|') enthalten."
        End If
    End With
End Sub




Public Function FindParagraphBeforeMustBe(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    Dim ParaIndex As Long
    Dim ParaBefore As Integer
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

    Dim ParaIndex As Long
    Dim ParaBefore As Integer
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
    ' MsgBox "Function FindParagraphAfterMustBe() wird gestartet mit Format " & NameOfFormat & " (" & ParaStyle & ")."
    Dim ParaIndex As Long
    Dim ParaAfter As Integer
    FindParagraphAfterMustBe = ParaIndex
    ' ParaAfter = ParaIndex + 1
    For ParaIndex = 1 To SearchRange.Paragraphs.Count
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
                If ActiveDocument.Paragraphs(ParaAfter).Range.Style = NameOfFormatAfter Then
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt korrekterweise Format " & NameOfFormatAfter
                    correctFormat = True
                Else
                    ' MsgBox "Auf Format " & NameOfFormat & " folgt nicht Format " & NameOfFormatAfter
                End If
            ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
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
    ' MsgBox "Function FindParagraphAfterMustNotBe() wird gestartet mit Format " & NameOfFormat & " (" & ParaStyle & ")."
    Dim ParaIndex As Long
    Dim ParaAfter As Integer
    FindParagraphAfterMustNotBe = ParaIndex
    ' ParaAfter = ParaIndex + 1
    For ParaIndex = 1 To SearchRange.Paragraphs.Count
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
            ' MsgBox "Durchlauf für " & NameOfFormatAfter & " beendet. correctFormat: " & correctFormat
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
