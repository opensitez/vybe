' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_reimplementation_in_derived_class
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

Interface IPrintable
    Sub Print()
End Interface

Class Parent
    Implements IPrintable
    Public Overridable Sub Print() Implements IPrintable.Print
        Console.WriteLine("Parent Print")
    End Sub
End Class

Class Child
    Inherits Parent
    Implements IPrintable
    Public Overrides Sub Print() Implements IPrintable.Print
        Console.WriteLine("Child Reimplemented Print")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IPrintable = New Child()
        p.Print()
    End Sub
End Module
