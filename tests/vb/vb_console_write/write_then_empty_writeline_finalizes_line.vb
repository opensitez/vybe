' vybe-test: vb/vb_console_write/write_then_empty_writeline_finalizes_line
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write("prefix")
        Console.WriteLine()
        Console.WriteLine("next")
    End Sub
End Module
