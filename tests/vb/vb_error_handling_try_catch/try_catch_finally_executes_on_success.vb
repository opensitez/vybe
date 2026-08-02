' vybe-test: vb/vb_error_handling_try_catch/try_catch_finally_executes_on_success
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Console.WriteLine("Try")
Catch
Console.WriteLine("Catch")
Finally
Console.WriteLine("Finally")
End Try
End Sub
End Module
