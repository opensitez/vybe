' vybe-test: vb/vb_error_handling_try_catch/try_catch_fallthrough_to_general
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Throw New System.InvalidOperationException()
Catch ex As System.DivideByZeroException
Console.WriteLine("DivByZero")
Catch ex As System.Exception
Console.WriteLine("General")
End Try
End Sub
End Module
