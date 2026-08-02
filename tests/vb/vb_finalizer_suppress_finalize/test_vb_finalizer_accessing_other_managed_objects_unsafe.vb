' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_accessing_other_managed_objects_unsafe
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

Imports System

Class ManagedDependency
    Public Sub Ping()
        Console.WriteLine("Dependency Ping")
    End Sub
End Class

Class Consumer
    Private dep As New ManagedDependency()

    Protected Overrides Sub Finalize()
        ' Note: Accessing dep in Finalize is unsafe because dep may already be finalized!
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim c As New Consumer()
        End Sub()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Finalization Finished")
    End Sub
End Module
