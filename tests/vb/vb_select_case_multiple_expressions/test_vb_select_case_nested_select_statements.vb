' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_nested_select_statements
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim category = "Tech"
        Dim subCat = "Mobile"
        Select Case category
            Case "Tech"
                Select Case subCat
                    Case "Mobile"
                        Console.WriteLine("Tech-Mobile")
                    Case "Desktop"
                        Console.WriteLine("Tech-Desktop")
                End Select
        End Select
    End Sub
End Module
