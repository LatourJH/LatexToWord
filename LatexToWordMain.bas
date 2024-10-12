Option Explicit

Sub MainSequence()
    On Error GoTo ErrorHandler
    
    ' === Step 1: Normalize Text ===
    Call NormalizeDocumentText
    Debug.Print "Step 1 Complete: Document text normalized."
    
    ' === Step 2: Convert LaTeX Equations with Python (Pass Document Path) ===
    Call ConvertLatexWithPython(ActiveDocument.FullName)
    Debug.Print "Step 2 Complete: LaTeX equations sent to Python."
    
    ' === Step 3: Read and Insert Processed LaTeX/MathML ===
    Dim outputBuffer As String
    outputBuffer = ReadPythonOutput() ' Read Python output from temp file
    If outputBuffer <> "" Then
        Call ReadAndInsertMathML(outputBuffer)
        Debug.Print "Step 3 Complete: Processed LaTeX equations inserted."
    Else
        Debug.Print "Step 3 Failed: No LaTeX equations were returned from Python."
    End If
    
    ' === Step 4: Apply Professional Format to all OMath objects ===
    Call ConvertAllToProfessional
    Debug.Print "Step 4 Complete: Converted all equations to Professional format."
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub NormalizeDocumentText()
    Dim textRange As Range
    Set textRange = ActiveDocument.Content
    
    ' Clear any font formatting to make sure all text uses the default font
    textRange.Font.Reset
    textRange.ParagraphFormat.Reset
End Sub

Sub ConvertLatexWithPython(ByVal docPath As String)
    Dim pythonScript As String
    Dim shellCommand As String
    Dim wsh As Object
    Dim tempFile As String
    
    ' Path to your Python script
    pythonScript = "C:\Users\latou\Desktop\LatexToWordProject\PythonToLatexMainFile.py"
    
    ' Temp file to store Python output
    tempFile = "C:\Users\latou\Desktop\LatexToWordProject\latex_output.txt"
    
    ' Construct the shell command to pass document path and temp output file
    shellCommand = "python """ & pythonScript & """ """ & docPath & """ """ & tempFile & """" 
    
    ' Use WScript.Shell for better handling of shell commands
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run shellCommand, 0, True ' Wait for the command to complete (3rd argument True)
End Sub

Function ReadPythonOutput() As String
    Dim fso As Object
    Dim tempFile As String
    Dim fileStream As Object
    Dim pythonOutput As String
    
    ' Temp file where Python wrote the output
    tempFile = "C:\Users\latou\Desktop\LatexToWordProject\latex_output.txt"
    
    ' Check if the file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tempFile) Then
        ' Open the file and read its contents
        Set fileStream = fso.OpenTextFile(tempFile, 1)
        pythonOutput = fileStream.ReadAll
        fileStream.Close
        
        ' Remove the temp file after reading
        fso.DeleteFile tempFile
    Else
        pythonOutput = "" ' No output from Python
    End If
    
    ReadPythonOutput = pythonOutput
End Function

Sub ReadAndInsertMathML(ByVal outputBuffer As String)
    Dim eqRange As Range
    Set eqRange = ActiveDocument.Content
    eqRange.Collapse Direction:=wdCollapseEnd
    
    ' Split the output into individual lines (one equation per line)
    Dim equations() As String
    equations = Split(outputBuffer, vbCrLf) ' Split by line break
    
    ' Insert each equation as an OMath object
    Dim equation As Variant
    For Each equation In equations
        If Len(Trim(equation)) > 0 Then ' Skip empty lines
            eqRange.InsertAfter equation
            eqRange.InsertParagraphAfter
            
            ' Convert inserted equation to OMath and build up
            eqRange.OMaths.Add eqRange
            eqRange.OMaths(1).BuildUp
            eqRange.Collapse Direction:=wdCollapseEnd
        End If
    Next equation
End Sub

Sub ConvertAllToProfessional()
    ' Convert all OMath objects in the document to Professional format
    Dim omath As omath
    For Each omath In ActiveDocument.OMaths
        ' Apply the professional format to the equation
        omath.BuildUp
    Next omath
End Sub
