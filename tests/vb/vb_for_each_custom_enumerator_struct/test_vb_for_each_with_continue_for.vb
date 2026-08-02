' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_with_continue_for
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        Dim oddSum = 0
        For Each n In numbers
            If n Mod 2 = 0 Then Continue For
            oddSum += n
        Next
        Console.WriteLine(oddSum)
    End Sub
End Module
