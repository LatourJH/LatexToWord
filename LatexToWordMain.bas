Attribute VB_Name = "LatexToWordMain"
Option Explicit

Sub MainSequence()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== VBA Macro Execution Started ==="
    
    ' === Step 1: Normalize Text ===
    Debug.Print "Step 1: Starting text normalization."
    Call NormalizeDocumentText
    Debug.Print "Step 1 Complete: Document text normalized."
    
    ' === Step 2: Convert LaTeX Equations with Python (Pass Document Path) ===
    Dim docPath As String
    docPath = ActiveDocument.FullName
    Debug.Print "Step 2: Converting LaTeX equations using Python."
    Debug.Print "Document Path: " & docPath
    Call ConvertLatexWithPython(docPath)
    Debug.Print "Step 2 Complete: LaTeX equations sent to Python."
    
    ' === Step 3: Read and Insert Processed LaTeX/MathML ===
    Dim outputBuffer As String
    Debug.Print "Step 3: Reading Python output from output file."
    outputBuffer = ReadPythonOutput() ' Read Python output from output file
    If outputBuffer <> "" Then
        Debug.Print "Step 3: Inserting processed LaTeX equations."
        Call ReadAndInsertMathML(outputBuffer)
        Debug.Print "Step 3 Complete: Processed LaTeX equations inserted."
    Else
        Debug.Print "Step 3 Failed: No LaTeX equations were returned from Python."
    End If
    
    ' === Step 4: Apply Professional Format to all OMath objects ===
    Debug.Print "Step 4: Converting all OMath objects to Professional format."
    Call ConvertAllToProfessional
    Debug.Print "Step 4 Complete: Converted all equations to Professional format."
    
    Debug.Print "=== VBA Macro Execution Completed ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error occurred: " & Err.Description
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub NormalizeDocumentText()
    Dim textRange As Range
    Set textRange = ActiveDocument.Content
    
    Debug.Print "Normalizing document text."
    
    ' Clear any font formatting to make sure all text uses the default font
    textRange.Font.Reset
    textRange.ParagraphFormat.Reset
End Sub

Sub ConvertLatexWithPython(ByVal docPath As String)
    Dim pythonScript As String
    Dim shellCommand As String
    Dim wsh As Object
    
    ' Path to your Python script
    pythonScript = "C:\Users\latou\Desktop\LatexToWordProject\PythonToLatexMainFile.py"
    Debug.Print "Python Script Path: " & pythonScript
    
    ' Construct the shell command to pass document path
    shellCommand = "python """ & pythonScript & """ """ & docPath & """"
    Debug.Print "Shell Command: " & shellCommand
    
    ' Use WScript.Shell for better handling of shell commands
    Set wsh = CreateObject("WScript.Shell")
    Debug.Print "Executing Python script..."
    wsh.Run shellCommand, 0, True ' Wait for the command to complete (3rd argument True)
    Debug.Print "Python script execution completed."
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
        Debug.Print "Temp file found: " & tempFile
        
        ' Open the file and read its contents with UTF-8 encoding
        With CreateObject("ADODB.Stream")
            .Type = 2 ' Specify stream type as text
            .Charset = "utf-8" ' Specify the encoding as UTF-8
            .Open
            .LoadFromFile tempFile
            pythonOutput = .ReadText ' Read the text as UTF-8
            .Close
        End With
        
        Debug.Print "Python output read successfully."
        Debug.Print "Python Output: " & pythonOutput
        
        ' Remove the temp file after reading
        fso.DeleteFile tempFile
        Debug.Print "Temp file deleted."
    Else
        Debug.Print "Temp file not found: " & tempFile
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
    
    Debug.Print "Inserting equations into the document."
    Debug.Print "Number of equations: " & UBound(equations) + 1
    
    ' Insert each equation as an OMath object
    Dim equation As Variant
    For Each equation In equations
        If Len(Trim(equation)) > 0 Then ' Skip empty lines
            Debug.Print "Inserting equation: " & equation
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
    Debug.Print "Converting all OMath objects to Professional format."
    For Each omath In ActiveDocument.OMaths
        ' Apply the professional format to the equation
        omath.BuildUp
    Next omath
End Sub

