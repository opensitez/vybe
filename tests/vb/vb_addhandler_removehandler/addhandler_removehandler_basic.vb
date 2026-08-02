' vybe-test: vb/vb_addhandler_removehandler/addhandler_removehandler_basic
' origin: languages/vb/tests/vb/test_vb_addhandler_removehandler.rs

Class Button
    Public Event Click()
    
    Public Sub PerformClick()
        RaiseEvent Click()
    End Sub
End Class

Module M
    Sub OnClick1()
        Console.WriteLine("Click 1")
    End Sub
    
    Sub OnClick2()
        Console.WriteLine("Click 2")
    End Sub

    Sub Main()
        Dim btn As New Button()
        
        AddHandler btn.Click, AddressOf OnClick1
        AddHandler btn.Click, AddressOf OnClick2
        btn.PerformClick()
        
        RemoveHandler btn.Click, AddressOf OnClick1
        btn.PerformClick()
    End Sub
End Module
