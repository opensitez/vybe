' vybe-test: vb/vb_oop_edges/notoverridable_method_in_child
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

Class Base
    Public Overridable Sub Run()
        Console.WriteLine("Base")
    End Sub
End Class

Class Child
    Inherits Base
    Public NotOverridable Overrides Sub Run()
        Console.WriteLine("Child")
    End Sub
End Class

Class GrandChild
    Inherits Child
    ' Cannot override Run here
End Class

Module M
    Sub Main()
        Dim c As New GrandChild()
        c.Run()
    End Sub
End Module
