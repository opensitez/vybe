' vybe-test: vb/vb_comprehensive/logical_orelse_short_circuit
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        If True OrElse False Then
            Console.WriteLine("yes")
        Else
            Console.WriteLine("no")
        End If
    End Sub
End Module
