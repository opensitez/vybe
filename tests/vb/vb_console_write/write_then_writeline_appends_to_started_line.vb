' vybe-test: vb/vb_console_write/write_then_writeline_appends_to_started_line
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write("x")
        Console.WriteLine("y")
    End Sub
End Module
