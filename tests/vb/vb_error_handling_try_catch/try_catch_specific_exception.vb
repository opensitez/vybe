' vybe-test: vb/vb_error_handling_try_catch/try_catch_specific_exception
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Dim x = 1 \ 0
Catch ex As System.DivideByZeroException
Console.WriteLine("DivByZero")
Catch ex As System.Exception
Console.WriteLine("General")
End Try
End Sub
End Module
