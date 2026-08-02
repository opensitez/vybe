' vybe-test: vb/vb_spec_error_handling_resources/error_spec_on_error_goto_label_jumps_to_handler
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        On Error GoTo Handler
        Err.Raise(5)
        Console.WriteLine("after")
        Exit Sub
Handler:
        Console.WriteLine("handled")
    End Sub
End Module
