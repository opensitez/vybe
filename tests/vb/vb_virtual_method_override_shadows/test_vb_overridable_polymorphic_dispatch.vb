' vybe-test: vb/vb_virtual_method_override_shadows/test_vb_overridable_polymorphic_dispatch
' origin: languages/vb/tests/vb/test_vb_virtual_method_override_shadows.rs

Class Animal
    Public Overridable Sub Speak()
        Console.WriteLine("Animal sound")
    End Sub
End Class

Class Dog
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("Woof")
    End Sub
End Class

Module Program
    Sub Main()
        Dim a As Animal = New Dog()
        a.Speak()
    End Sub
End Module
