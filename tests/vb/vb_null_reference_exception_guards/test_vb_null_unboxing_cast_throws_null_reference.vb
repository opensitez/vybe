' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_unboxing_cast_throws_null_reference
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

Imports System

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Try
            Dim i As Integer = CInt(obj)
            Console.WriteLine(i)
        Catch ex As NullReferenceException
            Console.WriteLine("Unboxing Null NullReferenceException Caught")
        Catch ex As Exception
            Console.WriteLine("Caught: " & ex.GetType().Name)
        End Try
    End Sub
End Module
