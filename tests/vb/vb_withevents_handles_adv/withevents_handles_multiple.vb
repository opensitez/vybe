' vybe-test: vb/vb_withevents_handles_adv/withevents_handles_multiple
' origin: languages/vb/tests/vb/test_vb_withevents_handles_adv.rs

Class Button
    Public Event Click()
    Public Sub PerformClick()
        RaiseEvent Click()
    End Sub
End Class

Class Form
    Private WithEvents btn1 As New Button()
    Private WithEvents btn2 As New Button()
    
    ' Handles multiple events
    Private Sub CommonClickHandler() Handles btn1.Click, btn2.Click
        Console.WriteLine("Clicked")
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
