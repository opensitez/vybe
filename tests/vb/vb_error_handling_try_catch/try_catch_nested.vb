' vybe-test: vb/vb_error_handling_try_catch/try_catch_nested
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Try
Throw New System.Exception()
Catch
Console.WriteLine("Inner")
End Try
Catch
Console.WriteLine("Outer")
End Try
End Sub
End Module
