'----------------------------------------------------------
'----- SET GLOBAL VARIABLES -----
'----------------------------------------------------------
Option Explicit
Public NameOfFormat As String
Public NameOfFormatAfter As String
Public NameOfFormatBefore As String
Dim logFile As Object
Dim logFilePath As String
Dim logFileName As String


Sub check_word_file()

'----------------------------------------------------------
'----- DELETE LOG FILE, IF EXISTS -----
'----------------------------------------------------------

' Set name of log file
logFileName = "log" & ".txt"
 
' Set the path for the log file
logFilePath = ActiveDocument.Path & "/" & logFileName

If Dir(logFilePath) <> "" Then
    'MsgBox "Log-Datei besteht bereits. Wird gelöscht."
    Kill logFilePath
End If

'----------------------------------------------------------
'----- SET HEADER FORMATS -----
'----------------------------------------------------------

' SET FORMAT OF FIRST PARAGRAPH
Paragraphs.First.Range.Select
search_firstPara
'Format will be set in Sub if check is TRUE

' SET FORMAT OF SECOND PARAGRAPH
Paragraphs(2).Style = "SuS_Headline"

' SET FORMAT OF THIRD PARAGRAPH
Paragraphs(3).Style = "SuS_Subhead1"


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
NameOfFormatAfter = "SuS_Autorname"
IsFound = FindParagraphAfter(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Bilddateiname"
NameOfFormatAfter = "SuS_Bild/Tabellenunterschrift"
IsFound = FindParagraphAfter(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)

NameOfFormat = "SuS_Bild/Tabellenunterschrift"
NameOfFormatAfter = "SuS_Mengentext"
IsFound = FindParagraphAfter(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)


'----------------------------------------------------------
'----- CHECK WHETHER FORMAT X EXISTS AND IF SO, CHECK WHETHER PREVIOUS PARAGRAPH IS FORMAT Y -----
'----------------------------------------------------------
NameOfFormat = "SuS_Bilddateiname"
NameOfFormatBefore = "SuS_Mengentext"
IsFound = FindParagraphBefore(ActiveDocument.StoryRanges(wdMainTextStory), NameOfFormat)


'----------------------------------------------------------
'----- CHECK IF NUMBER OF INSTANCES FORMAT X IS AN INTEGER MULTIPLE OF 2 (FORMAT ALWAYS NEEDS TO BE CLOSED) -----
'----------------------------------------------------------

' MOD Dividiert zwei Zahlen und gibt nur den Rest zurück.
'Result = number1 Mod number2
' Wenn Rest gleich 0 dann ist es ein ganzzahliges Vielfaches von 2 – wird also geöffnet UND geschlossen

'----------------------------------------------------------
'----- END OF SCRIPT MESSAGE -----
'----------------------------------------------------------

MsgBox "Das Skript ist fertig."

End Sub

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
            Paragraphs.First.Style = "SuS_Mengentext"
        Else
            WriteLogFile "Fehler in Zeile 1: " & " Absatz muss eine Pipe ('|') enthalten."
        End If
    End With
End Sub




Public Function FindParagraphBefore(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    Dim ParaIndex As Long
    Dim ParaBefore As Integer
    For ParaIndex = 1 To SearchRange.Paragraphs.Count

        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then
            'MsgBox "Absatzformat " & NameOfFormat & " gefunden in Zeile( " & ParaIndex & ")"
            FindParagraphBefore = ParaIndex
            'jump 1 paragraph back and check if it has certain format
            ParaBefore = ParaIndex - 1
            If ActiveDocument.Paragraphs(ParaBefore).Range.Style = NameOfFormatBefore Then
                'MsgBox "Alles OK in Zeile " & ParaIndex
            Else
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Absatz vor Absatzformat " & ParaStyle & " muss stets Absatzformat " & NameOfFormatAfter & " sein."
                'MsgBox "Fehler in Zeile " & ParaIndex & vbCrLf & "Absatz vor Absatzformat " & ParaStyle & " muss stets Absatzformat " & NameOfFormatAfter & " sein."
            End If
            
        Else
            'MsgBox "Absatzformat " & NameOfFormat & " wurde nicht gefunden."

        End If

    Next

End Function



Public Function FindParagraphAfter(ByVal SearchRange As Word.Range, ByVal ParaStyle As String) As Long

    Dim ParaIndex As Long
    Dim ParaAfter As Integer
    For ParaIndex = 1 To SearchRange.Paragraphs.Count

        If ActiveDocument.Paragraphs(ParaIndex).Range.Style = ParaStyle Then

            FindParagraphAfter = ParaIndex
            'jump 1 paragraph ahaed and check if it has certain format
            ParaAfter = ParaIndex + 1
            If ActiveDocument.Paragraphs(ParaAfter).Range.Style = NameOfFormatAfter Then
                'MsgBox "Passt"
            Else
                WriteLogFile "Fehler in Zeile " & ParaIndex & ": Auf Absatzformat " & ParaStyle & " muss stets Absatzformat " & NameOfFormatAfter & " folgen."
                'MsgBox "Fehler in Zeile " & ParaIndex & vbCrLf & "Auf Absatzformat " & ParaStyle & " muss stets Absatzformat " & NameOfFormatAfter & " folgen."
            End If

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