' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_nested_do_loops_exit_inner
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim outerCount = 0
        Dim innerCount = 0
        Do While outerCount < 2
            outerCount += 1
            Do
                innerCount += 1
                If innerCount >= 2 Then Exit Do
            Loop While True
        Loop
        Console.WriteLine(outerCount & "|" & innerCount)
    End Sub
End Module
