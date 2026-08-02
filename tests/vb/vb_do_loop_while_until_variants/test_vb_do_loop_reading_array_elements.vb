' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_loop_reading_array_elements
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40}
        Dim idx = 0
        Dim sum = 0
        Do While idx < numbers.Length
            sum += numbers(idx)
            idx += 1
        Loop
        Console.WriteLine(sum)
    End Sub
End Module
