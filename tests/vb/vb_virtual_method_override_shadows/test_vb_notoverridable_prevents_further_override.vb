' vybe-test: vb/vb_virtual_method_override_shadows/test_vb_notoverridable_prevents_further_override
' origin: languages/vb/tests/vb/test_vb_virtual_method_override_shadows.rs

Class BaseClass
    Public Overridable Sub Action()
        Console.WriteLine("Base")
    End Sub
End Class

Class MidClass
    Inherits BaseClass
    Public NotOverridable Overrides Sub Action()
        Console.WriteLine("Mid")
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As BaseClass = New MidClass()
        b.Action()
    End Sub
End Module
