' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_with_exit_select
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim x = 1
        Select Case x
            Case 1
                Console.WriteLine("Start Case 1")
                If x = 1 Then Exit Select
                Console.WriteLine("End Case 1")
        End Select
        Console.WriteLine("After Select")
    End Sub
End Module
