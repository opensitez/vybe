' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_with_byref_mutation_inside
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Private Sub Increment(ByRef val As Integer)
        val += 1
    End Sub

    Sub Main()
        Dim count = 0
        Do
            Increment(count)
        Loop Until count = 4
        Console.WriteLine(count)
    End Sub
End Module
