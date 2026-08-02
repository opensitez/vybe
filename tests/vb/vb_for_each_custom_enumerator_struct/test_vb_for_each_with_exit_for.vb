' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_with_exit_for
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        Dim lastSeen = 0
        For Each n In numbers
            If n = 4 Then Exit For
            lastSeen = n
        Next
        Console.WriteLine(lastSeen)
    End Sub
End Module
