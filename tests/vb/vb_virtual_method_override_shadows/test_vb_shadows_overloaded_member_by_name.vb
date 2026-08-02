' vybe-test: vb/vb_virtual_method_override_shadows/test_vb_shadows_overloaded_member_by_name
' origin: languages/vb/tests/vb/test_vb_virtual_method_override_shadows.rs

Class BasePrinter
    Public Sub Print(x As Integer)
        Console.WriteLine("Base Int: " & x)
    End Sub
End Class

Class DerivedPrinter
    Inherits BasePrinter
    Public Shadows Sub Print(s As String)
        Console.WriteLine("Derived String: " & s)
    End Sub
End Class

Module Program
    Sub Main()
        Dim dp As New DerivedPrinter()
        dp.Print("Hello")
    End Sub
End Module
