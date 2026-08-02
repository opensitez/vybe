' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_while_wend_legacy_loop
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim i = 0
        While i < 3
            i += 1
        End While
        Console.WriteLine(i)
    End Sub
End Module
