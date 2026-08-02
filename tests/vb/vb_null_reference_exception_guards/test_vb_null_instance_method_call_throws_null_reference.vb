' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_instance_method_call_throws_null_reference
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

Imports System

Class Document
    Public Sub Print() : Console.WriteLine("Print") : End Sub
End Class

Module Program
    Sub Main()
        Dim doc As Document = Nothing
        Try
            doc.Print()
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on Method Call")
        End Try
    End Sub
End Module
