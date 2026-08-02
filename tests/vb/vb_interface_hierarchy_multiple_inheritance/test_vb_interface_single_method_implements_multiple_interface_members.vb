' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_single_method_implements_multiple_interface_members
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

Interface IX
    Sub Common()
End Interface

Interface IY
    Sub Common()
End Interface

Class SharedImpl
    Implements IX, IY
    Public Sub Common() Implements IX.Common, IY.Common
        Console.WriteLine("Shared Common Execution")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New SharedImpl()
        Dim x As IX = obj
        Dim y As IY = obj
        x.Common()
        y.Common()
    End Sub
End Module
