' vybe-test: vb/vb_select_case_compare_text/select_case_compare_text
' origin: languages/vb/tests/vb/test_vb_select_case_compare_text.rs

Option Compare Text

Module M
    Sub Main()
        Dim s = "hello"
        
        ' Select Case with Option Compare Text should be case insensitive
        Select Case s
            Case "HELLO"
                Console.WriteLine("Matched")
            Case Else
                Console.WriteLine("Not Matched")
        End Select
    End Sub
End Module
