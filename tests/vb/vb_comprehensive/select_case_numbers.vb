' vybe-test: vb/vb_comprehensive/select_case_numbers
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case 3
                Console.WriteLine("three")
        End Select
    End Sub
End Module
