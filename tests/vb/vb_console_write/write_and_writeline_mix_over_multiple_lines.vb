' vybe-test: vb/vb_console_write/write_and_writeline_mix_over_multiple_lines
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write("a")
        Console.Write("b")
        Console.WriteLine("c")
        Console.WriteLine("d")
    End Sub
End Module
