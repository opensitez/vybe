' vybe-test: vb/vb_console_write/write_bool_then_text_keeps_same_line
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write(False)
        Console.Write("!")
        Console.WriteLine()
    End Sub
End Module
