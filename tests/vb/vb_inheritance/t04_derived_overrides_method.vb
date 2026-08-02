' vybe-test: vb/vb_inheritance/t04_derived_overrides_method
' origin: languages/vb/tests/vb/vb_inheritance_test.rs

Class Animal
    Sub Speak()
        Console.WriteLine("generic")
    End Sub
End Class

Class Dog
    Inherits Animal

    Sub Speak()
        Console.WriteLine("woof")
    End Sub
End Class

Dim d As New Dog()
d.Speak()
