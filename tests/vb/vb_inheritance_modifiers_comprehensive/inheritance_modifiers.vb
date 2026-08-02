' vybe-test: vb/vb_inheritance_modifiers_comprehensive/inheritance_modifiers
' origin: languages/vb/tests/vb/test_vb_inheritance_modifiers_comprehensive.rs

Class Base
    Public Overridable Sub Show()
        Console.WriteLine("Base")
    End Sub
End Class

Class Middle
    Inherits Base
    ' NotOverridable seals the method from further overriding
    Public NotOverridable Overrides Sub Show()
        Console.WriteLine("Middle")
    End Sub
End Class

Class Bottom
    Inherits Middle
    ' Cannot override Show here because it is NotOverridable in Middle
    Public Shadows Sub Show()
        Console.WriteLine("Bottom")
    End Sub
End Class

Module M
    Sub Main()
        Dim b As Base = New Bottom()
        ' Calls the sealed override in Middle
        b.Show()
        
        Dim bot As Bottom = New Bottom()
        bot.Show()
    End Sub
End Module
