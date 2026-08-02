' vybe-test: vb/vb_comprehensive/logical_andalso_short_circuit
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim x As Integer = 0
        If False AndAlso (x = 1) Then
            Console.WriteLine("yes")
        Else
            Console.WriteLine("no")
        End If
    End Sub
End Module
