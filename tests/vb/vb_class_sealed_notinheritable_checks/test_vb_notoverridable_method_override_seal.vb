' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_method_override_seal
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

Class BaseClass
    Public Overridable Sub Display()
        Console.WriteLine("Base Display")
    End Sub
End Class

Class MiddleClass
    Inherits BaseClass
    Public NotOverridable Overrides Sub Display()
        Console.WriteLine("Middle Sealed Display")
    End Sub
End Class

Module Program
    Sub Main()
        Dim m As BaseClass = New MiddleClass()
        m.Display()
    End Sub
End Module
