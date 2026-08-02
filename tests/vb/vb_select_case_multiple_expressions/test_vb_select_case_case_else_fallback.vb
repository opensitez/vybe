' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_case_else_fallback
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim color = "Purple"
        Select Case color
            Case "Red"
                Console.WriteLine("R")
            Case "Blue"
                Console.WriteLine("B")
            Case Else
                Console.WriteLine("Fallback")
        End Select
    End Sub
End Module
