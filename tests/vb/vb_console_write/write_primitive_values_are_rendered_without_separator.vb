' vybe-test: vb/vb_console_write/write_primitive_values_are_rendered_without_separator
' origin: languages/vb/tests/vb/test_vb_console_write.rs

Module M
    Sub Main()
        Console.Write(1)
        Console.Write(2)
        Console.Write(3)
        Console.WriteLine()
    End Sub
End Module
