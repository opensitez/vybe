' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_instance_property_getter_throws_null_reference
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

Imports System

Class User
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim u As User = Nothing
        Try
            Dim n = u.Name
            Console.WriteLine(n)
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on Property Get")
        End Try
    End Sub
End Module
