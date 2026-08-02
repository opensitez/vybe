' vybe-test: vb/vb_comprehensive/boolean_expressions_in_conditions
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim x As Integer = 5
        Dim y As Integer = 10
        If x > 3 And y < 20 Then
            Console.WriteLine("both true")
        Else
            Console.WriteLine("not both")
        End If
    End Sub
End Module
