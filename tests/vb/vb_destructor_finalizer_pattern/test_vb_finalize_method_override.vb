' vybe-test: vb/vb_destructor_finalizer_pattern/test_vb_finalize_method_override
' origin: languages/vb/tests/vb/test_vb_destructor_finalizer_pattern.rs

Imports System

Class FinalizableClass
    Protected Overrides Sub Finalize()
        Try
            Console.WriteLine("Finalized")
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fc As New FinalizableClass()
        fc = Nothing
        Console.WriteLine("Done")
    End Sub
End Module
