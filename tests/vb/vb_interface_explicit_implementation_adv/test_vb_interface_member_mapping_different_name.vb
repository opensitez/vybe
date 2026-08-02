' vybe-test: vb/vb_interface_explicit_implementation_adv/test_vb_interface_member_mapping_different_name
' origin: languages/vb/tests/vb/test_vb_interface_explicit_implementation_adv.rs

Interface IWorker
    Sub DoWork()
End Interface

Class Employee
    Implements IWorker

    Public Sub ExecuteJob() Implements IWorker.DoWork
        Console.WriteLine("Executing Job")
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Employee()
        e.ExecuteJob()
        Dim w As IWorker = e
        w.DoWork()
    End Sub
End Module
