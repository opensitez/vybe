' vybe-test: vb/vb_system_console_matrix/console_write_via_out_and_err_channels
' origin: languages/vb/tests/vb/test_vb_system_console_matrix.rs

Imports System

Module M
    Sub Main()
        Console.Write("out:")
        Console.Error.Write("err")
        Console.WriteLine("done")
    End Sub
End Module
