Attribute VB_Name = "LatexToWordMain"
Option Explicit

Private Const SCRIPT_FILE_NAME As String = "PythonToLatexMainFile.py"

' Convert the currently selected LaTeX without changing formatting elsewhere in
' the document. Select one equation, or several equations separated by lines,
' before running this macro.
Public Sub MainSequence()
    On Error GoTo ErrorHandler

    Dim selectedRange As Range
    Set selectedRange = Selection.Range.Duplicate

    If selectedRange.Start = selectedRange.End Then
        MsgBox "Select the LaTeX equation text first, then run LatexToWord again.", _
               vbInformation, "LatexToWord"
        Exit Sub
    End If

    Dim pythonScript As String
    pythonScript = FindPythonScript()
    If Len(pythonScript) = 0 Then Exit Sub

    Dim tempBase As String
    Dim inputFile As String
    Dim outputFile As String
    Dim logFile As String
    tempBase = CreateTempBasePath()
    inputFile = tempBase & "-input.txt"
    outputFile = tempBase & "-output.txt"
    logFile = tempBase & ".log"

    WriteUtf8Text inputFile, selectedRange.Text

    Dim exitCode As Long
    exitCode = RunConverter(pythonScript, inputFile, outputFile, logFile)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 1000, "LatexToWord", _
                  "Python could not convert the selected text. Diagnostic log: " & logFile
    End If
    If Not FileExists(outputFile) Then
        Err.Raise vbObjectError + 1001, "LatexToWord", _
                  "Python finished without creating an output file. Diagnostic log: " & logFile
    End If

    Dim outputBuffer As String
    outputBuffer = ReadUtf8Text(outputFile)
    If Len(Trim$(outputBuffer)) = 0 Then
        Err.Raise vbObjectError + 1002, "LatexToWord", _
                  "No equations were returned. Diagnostic log: " & logFile
    End If

    Dim undoStarted As Boolean
    StartUndoRecord undoStarted

    Dim convertedCount As Long
    convertedCount = ReplaceSelectionWithEquations(selectedRange, outputBuffer)

    EndUndoRecord undoStarted
    CleanupFile inputFile
    CleanupFile outputFile
    CleanupFile logFile

    MsgBox CStr(convertedCount) & " equation(s) converted.", vbInformation, "LatexToWord"
    Exit Sub

ErrorHandler:
    Dim errorDescription As String
    errorDescription = Err.Description
    On Error Resume Next
    EndUndoRecord undoStarted
    CleanupFile inputFile
    CleanupFile outputFile
    On Error GoTo 0

    MsgBox "LatexToWord could not complete the conversion." & vbCrLf & vbCrLf & _
           errorDescription, vbExclamation, "LatexToWord"
End Sub

Private Function FindPythonScript() As String
    Dim folder As Variant
    Dim candidate As String

    For Each folder In Array(ActiveDocument.Path, ThisDocument.Path)
        If Len(CStr(folder)) > 0 Then
            candidate = CStr(folder) & Application.PathSeparator & SCRIPT_FILE_NAME
            If FileExists(candidate) Then
                FindPythonScript = candidate
                Exit Function
            End If
        End If
    Next folder

    Dim picker As Object
    Set picker = Application.FileDialog(3) ' msoFileDialogFilePicker
    With picker
        .Title = "Locate " & SCRIPT_FILE_NAME
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Python files", "*.py"
        If .Show = -1 Then FindPythonScript = .SelectedItems(1)
    End With
End Function

Private Function RunConverter(ByVal pythonScript As String, _
                              ByVal inputFile As String, _
                              ByVal outputFile As String, _
                              ByVal logFile As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim projectFolder As String
    projectFolder = fso.GetParentFolderName(pythonScript)

    Dim virtualEnvPython As String
    virtualEnvPython = projectFolder & "\.venv\Scripts\python.exe"

    Dim pythonCommand As String
    If FileExists(virtualEnvPython) Then
        pythonCommand = QuoteArgument(virtualEnvPython)
    Else
        pythonCommand = "py -3"
    End If

    Dim shellCommand As String
    shellCommand = pythonCommand & " " & QuoteArgument(pythonScript) & _
                   " --text-file " & QuoteArgument(inputFile) & _
                   " --output " & QuoteArgument(outputFile) & _
                   " --log " & QuoteArgument(logFile) & _
                   " --selection-only"

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    RunConverter = shell.Run(shellCommand, 0, True)
End Function

Private Function ReplaceSelectionWithEquations(ByVal selectedRange As Range, _
                                               ByVal outputBuffer As String) As Long
    Dim normalizedOutput As String
    normalizedOutput = Replace(outputBuffer, vbCrLf, vbLf)
    normalizedOutput = Replace(normalizedOutput, vbCr, vbLf)

    Dim equations() As String
    equations = Split(normalizedOutput, vbLf)

    ' Preserve a final paragraph mark if Word included it in the selection.
    If selectedRange.End > selectedRange.Start Then
        If Right$(selectedRange.Text, 1) = vbCr Then
            selectedRange.MoveEnd Unit:=wdCharacter, Count:=-1
        End If
    End If

    selectedRange.Text = ""
    selectedRange.Collapse Direction:=wdCollapseStart

    Dim equation As Variant
    Dim equationText As String
    Dim mathRange As Range
    Dim equationStart As Long
    Dim count As Long

    For Each equation In equations
        equationText = Trim$(CStr(equation))
        If Len(equationText) > 0 Then
            If count > 0 Then
                selectedRange.InsertParagraphAfter
                selectedRange.Collapse Direction:=wdCollapseEnd
            End If

            equationStart = selectedRange.Start
            selectedRange.InsertAfter equationText
            Set mathRange = ActiveDocument.Range( _
                Start:=equationStart, End:=equationStart + Len(equationText))
            ActiveDocument.OMaths.Add mathRange
            mathRange.OMaths(1).BuildUp
            selectedRange.SetRange Start:=mathRange.End, End:=mathRange.End
            count = count + 1
        End If
    Next equation

    ReplaceSelectionWithEquations = count
End Function

Private Sub WriteUtf8Text(ByVal filePath As String, ByVal value As String)
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText value
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Function ReadUtf8Text(ByVal filePath As String) As String
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With
End Function

Private Function CreateTempBasePath() As String
    Randomize
    CreateTempBasePath = Environ$("TEMP") & Application.PathSeparator & _
                         "latextoword-" & Format$(Now, "yyyymmdd-hhnnss") & _
                         "-" & Format$(CLng(Rnd() * 1000000), "000000")
End Function

Private Function QuoteArgument(ByVal value As String) As String
    QuoteArgument = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem)) > 0)
End Function

Private Sub CleanupFile(ByVal filePath As String)
    If Len(filePath) = 0 Then Exit Sub
    If FileExists(filePath) Then Kill filePath
End Sub

Private Sub StartUndoRecord(ByRef started As Boolean)
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Convert LaTeX equations"
    started = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EndUndoRecord(ByRef started As Boolean)
    If Not started Then Exit Sub
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    started = False
    Err.Clear
    On Error GoTo 0
End Sub
