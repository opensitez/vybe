' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_char_ranges
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim ch As Char = "k"c
        Select Case ch
            Case "a"c To "m"c
                Console.WriteLine("First Half")
            Case "n"c To "z"c
                Console.WriteLine("Second Half")
        End Select
    End Sub
End Module
