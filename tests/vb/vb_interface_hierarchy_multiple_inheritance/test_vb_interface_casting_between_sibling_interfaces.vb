' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_casting_between_sibling_interfaces
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

Interface ISource
    Sub SourceAction()
End Interface

Interface ITarget
    Sub TargetAction()
End Interface

Class DualAction
    Implements ISource, ITarget
    Public Sub SourceAction() Implements ISource.SourceAction
        Console.WriteLine("Source")
    End Sub
    Public Sub TargetAction() Implements ITarget.TargetAction
        Console.WriteLine("Target")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As ISource = New DualAction()
        Dim t As ITarget = CType(s, ITarget)
        t.TargetAction()
    End Sub
End Module
