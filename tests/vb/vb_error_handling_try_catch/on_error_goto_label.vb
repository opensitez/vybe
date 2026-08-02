' vybe-test: vb/vb_error_handling_try_catch/on_error_goto_label
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
On Error GoTo ErrorHandler
Dim x = 1 \ 0
Console.WriteLine("Skipped")
Exit Sub
ErrorHandler:
Console.WriteLine("Handled")
End Sub
End Module
