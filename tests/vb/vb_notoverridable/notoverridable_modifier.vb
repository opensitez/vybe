' vybe-test: vb/vb_notoverridable/notoverridable_modifier
' origin: languages/vb/tests/vb/test_vb_notoverridable.rs

Class Base
    Public Overridable Sub Print()
        Console.WriteLine("Base")
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' NotOverridable seals the method from further overriding
    Public NotOverridable Overrides Sub Print()
        Console.WriteLine("Derived")
    End Sub
End Class

Class MoreDerived
    Inherits Derived
    ' Cannot override Print here, it's a compile error, but we just verify it runs Base/Derived
    Public Shadows Sub Print()
        Console.WriteLine("Shadowed")
    End Sub
End Class

Module M
    Sub Main()
        Dim md As New MoreDerived()
        md.Print()
        
        Dim b As Base = md
        b.Print()
    End Sub
End Module
