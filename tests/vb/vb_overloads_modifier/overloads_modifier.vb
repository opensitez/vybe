' vybe-test: vb/vb_overloads_modifier/overloads_modifier
' origin: languages/vb/tests/vb/test_vb_overloads_modifier.rs

Class Base
    Public Overridable Sub Show()
        Console.WriteLine("Base.Show")
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Overrides replaces the base implementation
    Public Overrides Sub Show()
        Console.WriteLine("Derived.Show")
    End Sub
    
    ' Overloads allows creating a new method with the same name but different signature
    Public Overloads Sub Show(message As String)
        Console.WriteLine("Derived.Show: " & message)
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.Show()
        d.Show("Hello")
    End Sub
End Module
