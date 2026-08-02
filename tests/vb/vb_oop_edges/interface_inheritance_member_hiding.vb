' vybe-test: vb/vb_oop_edges/interface_inheritance_member_hiding
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

Interface IBase
    Sub Test()
End Interface

Interface IDerived
    Inherits IBase
    Shadows Sub Test()
End Interface

Class C
    Implements IDerived
    
    Public Sub TestBase() Implements IBase.Test
        Console.WriteLine("Base")
    End Sub
    
    Public Sub TestDerived() Implements IDerived.Test
        Console.WriteLine("Derived")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As IDerived = New C()
        d.Test()
    End Sub
End Module
