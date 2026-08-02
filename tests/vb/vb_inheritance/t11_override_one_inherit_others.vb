' vybe-test: vb/vb_inheritance/t11_override_one_inherit_others
' origin: languages/vb/tests/vb/vb_inheritance_test.rs

Class Base
    Sub A()
        Console.WriteLine("base-A")
    End Sub

    Sub B()
        Console.WriteLine("base-B")
    End Sub
End Class

Class Child
    Inherits Base

    Sub A()
        Console.WriteLine("child-A")
    End Sub
End Class

Dim c As New Child()
c.A()
c.B()
