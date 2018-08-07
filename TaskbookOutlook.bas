Attribute VB_Name = "TaskbookOutlook"
Sub SendTask(boardName As String, text As String)
    Dim command As String
    command = "tb.cmd -t " + "@" + boardName + " """ + text + """"
    Shell (command)
End Sub

Function GetCurrentText() As String
    Dim item As Object
    Set item = Application.ActiveExplorer.selection(1)
    GetCurrentText = item.Body
End Function

Function GetBoardName() As String
    Dim item As Outlook.MailItem
    Set item = Application.ActiveExplorer.selection(1)
    GetBoardName = Replace(item.Sender, " ", "_")
End Function

Sub CreateTask()
    Dim currentText As String
    Dim regExp As New regExp
    
    Dim pattern As String: pattern = "From:.*<([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})>.*"
    Dim bodyText As String: bodyText = GetCurrentText()
    
    With regExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    currentText = regExp.Replace(Replace(bodyText, vbLf, " "), "::END SECTION::")
    
    Call SendTask(GetBoardName(), currentText)
End Sub
