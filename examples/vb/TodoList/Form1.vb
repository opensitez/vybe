Public Class Form1
    Dim taskCount As Integer
    Dim tasks(20) As String

    Private Sub Form1_Load() Handles Me.Load
        taskCount = 0
        UpdateStatus()
    End Sub

    Private Sub btnAdd_Click() Handles btnAdd.Click
        Dim newTask As String
        newTask = txtTask.Text

        If newTask = "" Then
            MsgBox("Please enter a task!")
            Exit Sub
        End If

        ' Add task using With block to update multiple controls
        With lblTitle
            .Text = "Todo List - Adding..."
        End With

        tasks(taskCount) = newTask
        taskCount = taskCount + 1
        
        ' Clear input
        txtTask.Text = ""

        ' Refresh the list display
        RefreshList()
        UpdateStatus()

        With lblTitle
            .Text = "Todo List"
        End With
    End Sub

    Private Sub btnRemove_Click() Handles btnRemove.Click
        If taskCount = 0 Then
            MsgBox("No tasks to remove!")
            Exit Sub
        End If

        ' Remove last task
        taskCount = taskCount - 1
        tasks(taskCount) = ""

        RefreshList()
        UpdateStatus()
    End Sub

    Private Sub btnClear_Click() Handles btnClear.Click
        If taskCount = 0 Then
            Exit Sub
        End If

        Dim i As Integer
        For i = 0 To taskCount - 1
            tasks(i) = ""
        Next

        taskCount = 0
        RefreshList()
        UpdateStatus()

        With lblTitle
            .Text = "Todo List - Cleared!"
        End With
    End Sub

    Private Sub RefreshList()
        ' Build display string from tasks array
        Dim display As String
        display = ""
        Dim i As Integer
        For i = 0 To taskCount - 1
            If i > 0 Then display = display & "|"
            display = display & (i + 1) & ". " & tasks(i)
        Next

        lstTasks.Text = display
    End Sub

    Private Sub UpdateStatus()
        With lblCount
            If taskCount = 0 Then
                .Text = "No tasks"
            ElseIf taskCount = 1 Then
                .Text = "1 task"
            Else
                .Text = taskCount & " tasks"
            End If
        End With
    End Sub
End Class
