' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_suppress_finalize_called_multiple_times
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

Imports System

Class MultiSuppressObject
    Protected Overrides Sub Finalize()
        Console.WriteLine("Finalize Ran")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New MultiSuppressObject()
        GC.SuppressFinalize(obj)
        GC.SuppressFinalize(obj)
        GC.SuppressFinalize(obj)

        obj = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Suppression Multi Safe")
    End Sub
End Module
