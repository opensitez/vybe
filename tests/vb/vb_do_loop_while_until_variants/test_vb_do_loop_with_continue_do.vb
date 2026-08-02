' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_with_continue_do
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim sum = 0
        Dim i = 0
        Do While i < 5
            i += 1
            If i = 3 Then Continue Do
            sum += i
        Loop
        Console.WriteLine(sum)
    End Sub
End Module
