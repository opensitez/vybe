' vybe-test: vb/vb_comprehensive/select_case_with_else
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim x As Integer = 99
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case Else
                Console.WriteLine("other")
        End Select
    End Sub
End Module
