' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_boolean_true_pattern_matching
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim age = 25
        Dim isStudent = True
        Select Case True
            Case age < 18
                Console.WriteLine("Minor")
            Case age >= 18 AndAlso isStudent
                Console.WriteLine("Student Adult")
            Case Else
                Console.WriteLine("Adult")
        End Select
    End Sub
End Module
