' vybe-test: vb/vb_select_case_advanced/select_case_is_operator
' origin: languages/vb/tests/vb/test_vb_select_case_advanced.rs

Module M
    Sub Main()
        Dim age As Integer = 15
        Select Case age
            Case Is >= 18
                Console.WriteLine("Adult")
            Case Is < 13
                Console.WriteLine("Child")
            Case Else
                Console.WriteLine("Teenager")
        End Select
    End Sub
End Module
