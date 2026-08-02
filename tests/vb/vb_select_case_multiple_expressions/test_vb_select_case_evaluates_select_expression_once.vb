' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_evaluates_select_expression_once
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Private Function GetValue(ByRef count As Integer) As Integer
        count += 1
        Return 5
    End Function

    Sub Main()
        Dim evalCount = 0
        Select Case GetValue(evalCount)
            Case 1
                Console.WriteLine("One")
            Case 5
                Console.WriteLine("Five|Evals=" & evalCount)
        End Select
    End Sub
End Module
