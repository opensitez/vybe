' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_first_matching_case_executes_only
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim val = 10
        Select Case val
            Case 10
                Console.WriteLine("First 10")
            Case Is >= 10
                Console.WriteLine("Second >= 10")
        End Select
    End Sub
End Module
