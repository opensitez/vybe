' vybe-test: vb/vb_error_handling_try_catch/try_catch_exit_try
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Try
Console.WriteLine("Start")
Exit Try
Console.WriteLine("End")
Catch
Finally
Console.WriteLine("Finally")
End Try
End Sub
End Module
