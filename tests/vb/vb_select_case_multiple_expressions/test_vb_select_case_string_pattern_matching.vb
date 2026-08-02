' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_string_pattern_matching
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim fruit = "Banana"
        Select Case fruit
            Case "Apple", "Pear"
                Console.WriteLine("Pome Fruit")
            Case "Banana", "Mango", "Pineapple"
                Console.WriteLine("Tropical Fruit")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
