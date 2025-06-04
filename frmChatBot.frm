VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChatBot 
   Caption         =   "UserForm1"
   ClientHeight    =   5544
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7368
   OleObjectBlob   =   "frmChatBot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChatBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim knowledgeBase As Object

Private Sub UserForm_Initialize()
    Set knowledgeBase = CreateObject("Scripting.Dictionary")
    LoadKnowledgeBase
    txtAnswer.text = "Hello! Ask me something or teach me below."
End Sub

Private Sub LoadKnowledgeBase()
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
    MsgBox "Could not load knowledge.txt", vbCritical
End Sub

Private Sub btnAsk_Click()
    Dim question As String
    question = Trim(LCase(txtQuestion.text))
    
    If question = "" Then
        txtAnswer.text = "? Please enter a question."
        Exit Sub
    End If
    
    If knowledgeBase.exists(question) Then
        txtAnswer.text = knowledgeBase(question)
    Else
        txtAnswer.text = "I don't know that yet. Teach me below!"
    End If
    
    LogInteraction question, txtAnswer.text
End Sub

Private Sub btnTeach_Click()
    Dim newQ As String
    Dim newA As String
    
    newQ = Trim(txtNewQ.text)
    newA = Trim(txtNewA.text)
    
    If newQ = "" Or newA = "" Then
        MsgBox "Please fill both new phrase and bot reply.", vbExclamation
        Exit Sub
    End If
    
    AppendToKnowledgeFile newQ, newA
    knowledgeBase(Trim(LCase(newQ))) = newA
    
    MsgBox "? I've learned: " & newQ, vbInformation
    txtNewQ.text = ""
    txtNewA.text = ""
End Sub

Private Sub btnClear_Click()
    txtQuestion.text = ""
    txtAnswer.text = ""
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub AppendToKnowledgeFile(question As String, answer As String)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\knowledge.txt"
    
    Dim fileNum As Integer: fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, Trim(question) & " - " & Trim(answer)
    Close #fileNum
End Sub

Private Sub LogInteraction(question As String, response As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Log"
        ws.Range("A1:C1").Value = Array("Time", "Question", "Response")
    End If
    On Error GoTo 0
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = question
    ws.Cells(lastRow, 3).Value = response
End Sub

