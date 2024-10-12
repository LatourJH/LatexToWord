Attribute VB_Name = "Module1"
Sub MainSequence()
    On Error GoTo ErrorHandler
    
    ' === Step 1: Normalize Text ===
    Call NormalizeDocumentText
    Debug.Print "Step 1 Complete: Document text normalized."
    
    ' === Step 2: Convert LaTeX Equations with Python (Pass Document Path) ===
    Call ConvertLatexWithPython(ActiveDocument.FullName)
    Debug.Print "Step 2 Complete: LaTeX equations sent to Python."
    
    ' === Step 3: Read and Insert Processed LaTeX/MathML ===
    Call ReadAndInsertMathML
    Debug.Print "Step 3 Complete: Processed LaTeX equations inserted."

    ' === Step 4: Apply Professional Format to all OMath objects ===
    Call ConvertAllToProfessional
    Debug.Print "Step 4 Complete: Converted all equations to Professional format."
    
    ' === Final Step: Completion Message ===
    MsgBox "All steps completed successfully! LaTeX equations have been processed and converted."
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Sub NormalizeDocumentText()
    Dim textRange As Range
    
    ' Set the range to cover the entire document
    Set textRange = ActiveDocument.Content
    
    ' Clear any font formatting to make sure all text uses the default font
    textRange.Font.Reset
    textRange.ParagraphFormat.Reset
End Sub

Sub ConvertLatexWithPython(docPath As String)
    Dim pythonScript As String
    Dim shellCommand As String
    
    ' Update the path to your Python script
    pythonScript = "C:\Users\latou\Desktop\PythonToLatexMainFile.py"

    ' Construct shell command to pass the document path
    shellCommand = "python """ & pythonScript & """ """ & docPath & """"
    
    ' Run the Python script with the document path
    Shell shellCommand, vbHide
    
    ' Wait for Python script to complete
    Dim pauseTime As Single
    pauseTime = Timer + 2 ' Pause for 2 seconds
    Do While Timer < pauseTime
        DoEvents
    Loop
End Sub
Sub ReadAndInsertMathML()
    Dim line As String
    Dim stdOut As Object
    Dim proc As Object
    Dim eqRange As Range
    Dim outputBuffer As String
    
    ' Create WScript.Shell object to run the Python process
    Set proc = CreateObject("WScript.Shell").Exec("python C:\Users\latou\Desktop\PythonToLatexMainFile.py")
    
    ' Capture output from the Python script line by line
    Set stdOut = proc.stdOut
    Set eqRange = ActiveDocument.Content
    eqRange.Collapse Direction:=wdCollapseEnd
    
    ' Prepare to collect the full output
    outputBuffer = ""
    
    ' Read each line printed by the Python script
    Do While Not stdOut.AtEndOfStream
        line = stdOut.ReadLine
        If Len(Trim(line)) > 0 Then
            outputBuffer = outputBuffer & line & vbCrLf ' Collect all lines
        End If
    Loop
    
    ' Split the outputBuffer into individual lines (one equation per line)
    Dim equations() As String
    equations = Split(outputBuffer, vbCrLf) ' Split by line break
    
    ' Insert each equation into the Word document
    Dim equation As Variant
    For Each equation In equations
        If Len(Trim(equation)) > 0 Then ' Skip empty lines
            ' Insert the equation/result into the Word document
            eqRange.InsertAfter equation
            eqRange.InsertParagraphAfter
            
            ' Convert inserted line to OMath (math object) and format
            eqRange.OMaths.Add eqRange
            eqRange.OMaths(1).BuildUp
            
            ' Collapse the range to prepare for the next insertion
            eqRange.Collapse Direction:=wdCollapseEnd
        End If
    Next equation
    
    MsgBox "Step 3 Complete: Processed LaTeX equations inserted."
End Sub


Sub ConvertAllToProfessional()
    ' Convert all OMath objects in the document to Professional format
    Dim omath As omath
    For Each omath In ActiveDocument.OMaths
        ' Apply the professional format to the equation
        omath.BuildUp
    Next omath
End Sub

