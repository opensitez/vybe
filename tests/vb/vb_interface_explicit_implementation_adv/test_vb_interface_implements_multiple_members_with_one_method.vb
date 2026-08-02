' vybe-test: vb/vb_interface_explicit_implementation_adv/test_vb_interface_implements_multiple_members_with_one_method
' origin: languages/vb/tests/vb/test_vb_interface_explicit_implementation_adv.rs

Interface IFoo
    Sub Process()
End Interface

Interface IBar
    Sub Process()
End Interface

Class Processor
    Implements IFoo, IBar

    Public Sub CommonProcess() Implements IFoo.Process, IBar.Process
        Console.WriteLine("Common Process")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        Dim f As IFoo = p
        Dim b As IBar = p
        f.Process()
        b.Process()
    End Sub
End Module
