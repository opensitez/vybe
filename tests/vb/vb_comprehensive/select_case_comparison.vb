' vybe-test: vb/vb_comprehensive/select_case_comparison
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim score As Integer = 85
        If score >= 90 Then
            Console.WriteLine("A")
        ElseIf score >= 80 Then
            Console.WriteLine("B")
        ElseIf score >= 70 Then
            Console.WriteLine("C")
        Else
            Console.WriteLine("F")
        End If
    End Sub
End Module
