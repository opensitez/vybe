' vybe-test: vb/vb_like_operator_pattern_matching/test_vb_like_operator_unclosed_bracket_matches_literal
' origin: languages/vb/tests/vb/test_vb_like_operator_pattern_matching.rs

Module Program
    Sub Main()
        ' Unclosed bracket in Like pattern throws or matches literally depending on runtime
        Try
            Dim res = "A[" Like "A["
            Console.WriteLine(res)
        Catch ex As System.Exception
            Console.WriteLine("Like Pattern Syntax Exception Caught")
        End Try
    End Sub
End Module
