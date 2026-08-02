' vybe-test: vb/vb_error_handling/exception_in_function
' origin: languages/vb/tests/vb/test_vb_error_handling.rs

Module M
    Function Divide(a As Integer, b As Integer) As Integer
        If b = 0 Then
            Throw New Exception("Division by zero")
        End If
        Return a \ b
    End Function
    Sub Main()
        Try
            Console.WriteLine(Divide(10, 2))
            Console.WriteLine(Divide(10, 0))
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
