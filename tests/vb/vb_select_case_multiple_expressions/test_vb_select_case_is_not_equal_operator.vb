' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_is_not_equal_operator
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim status = 200
        Select Case status
            Case Is <> 200
                Console.WriteLine("Error Status")
            Case Else
                Console.WriteLine("OK Status")
        End Select
    End Sub
End Module
