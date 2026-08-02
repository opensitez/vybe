' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_shadows_combination
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

Class GrandParent
    Public Overridable Sub Show() : Console.WriteLine("GrandParent") : End Sub
End Class

Class Parent
    Inherits GrandParent
    Public Overrides Sub Show() : Console.WriteLine("Parent") : End Sub
End Class

Class Child
    Inherits Parent
    Public Shadows Sub Show() : Console.WriteLine("Child Shadow") : End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Child()
        c.Show()
        Dim p As Parent = c
        p.Show()
    End Sub
End Module
