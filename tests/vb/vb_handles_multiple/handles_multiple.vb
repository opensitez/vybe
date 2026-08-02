' vybe-test: vb/vb_handles_multiple/handles_multiple
' origin: languages/vb/tests/vb/test_vb_handles_multiple.rs

Class Button
    Public Event Click()
    
    Public Sub PerformClick()
        RaiseEvent Click()
    End Sub
End Class

Class Form
    Private WithEvents btn1 As New Button()
    Private WithEvents btn2 As New Button()
    
    ' Handles clause with multiple events
    Private Sub Buttons_Click() Handles btn1.Click, btn2.Click
        Console.WriteLine("Button clicked")
    End Sub
    
    Public Sub Test()
        btn1.PerformClick()
        btn2.PerformClick()
    End Sub
End Class

Module M
    Sub Main()
        Dim f As New Form()
        f.Test()
    End Sub
End Module
