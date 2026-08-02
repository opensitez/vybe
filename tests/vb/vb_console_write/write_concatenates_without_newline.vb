' vybe-test: vb/vb_console_write/write_concatenates_without_newline
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write("a")
        Console.Write("b")
        Console.WriteLine()
    End Sub
End Module
