' vybe-test: vb/vb_error_handling_try_catch/try_catch_multiple_exceptions_same_block
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Throw New System.DivideByZeroException()
Catch ex As System.DivideByZeroException
Console.WriteLine("Div")
Catch ex As System.OverflowException
Console.WriteLine("Over")
End Try
End Sub
End Module
