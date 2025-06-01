Attribute VB_Name = "Module1"
' === Module1 ===
Dim knowledgeBase As Object

Sub InitializeBot()
    Set knowledgeBase = CreateObject("Scripting.Dictionary")
    LoadKnowledgeBase
    MsgBox "?? VBA Bot is ready! Use StartChat to begin a full conversation.", vbInformation
End Sub

Sub LoadKnowledgeBase()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\knowledge.txt"
    
    Dim line As String
    Dim fileNum As Integer: fileNum = FreeFile
    
    On Error GoTo FileError
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        If InStr(line, " - ") > 0 Then
            Dim parts() As String
            parts = Split(line, " - ", 2)
            knowledgeBase(Trim(LCase(parts(0)))) = Trim(parts(1))
        End If
    Loop
    Close #fileNum
    Exit Sub

FileError:
    MsgBox "Missing or unreadable knowledge.txt. Please place it in the workbook folder.", vbCritical
End Sub

Sub StartChat()
    Dim userInput As String
    Dim response As String
    Call InitializeBot
    
    Do
        userInput = InputBox("?? Ask me something (type 'exit' to quit):", "VBA Chat Bot")
        If userInput = "" Then Exit Do
        If LCase(userInput) = "exit" Then Exit Do
        
        response = GenerateResponse(userInput)
        MsgBox response, vbInformation, "Bot"
        LogInteraction userInput, response
    Loop
End Sub

Function GenerateResponse(inputText As String) As String
    Dim cleaned As String
    cleaned = Trim(LCase(inputText))
    
    If knowledgeBase.exists(cleaned) Then
        GenerateResponse = knowledgeBase(cleaned)
    Else
        GenerateResponse = "? I don't know how to respond to that. Type 'teach' to help me learn."
        
        If LCase(inputText) = "teach" Then
            TeachBot
            GenerateResponse = "Thanks! I've learned something new. Ask again."
        End If
    End If
End Function

Sub TeachBot()
    Dim newQuestion As String
    Dim newAnswer As String
    
    newQuestion = InputBox("?? What phrase should I learn?")
    If newQuestion = "" Then Exit Sub
    
    newAnswer = InputBox("?? What should I say when someone says: '" & newQuestion & "'?")
    If newAnswer = "" Then Exit Sub
    
    AppendToKnowledgeFile newQuestion, newAnswer
    MsgBox "? Got it! I've learned something new.", vbInformation
End Sub

Sub AppendToKnowledgeFile(question As String, answer As String)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\knowledge.txt"
    
    Dim fileNum As Integer: fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, Trim(question) & " - " & Trim(answer)
    Close #fileNum
    
    knowledgeBase(Trim(LCase(question))) = Trim(answer)
End Sub

Sub LogInteraction(userInput As String, response As String)
    Dim logSheet As Worksheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("Log")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add
        logSheet.Name = "Log"
        logSheet.Range("A1:C1").Value = Array("Time", "User Input", "Bot Response")
    End If
    On Error GoTo 0
    
    Dim lastRow As Long
    lastRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    logSheet.Cells(lastRow, 1).Value = Now
    logSheet.Cells(lastRow, 2).Value = userInput
    logSheet.Cells(lastRow, 3).Value = response
End Sub

