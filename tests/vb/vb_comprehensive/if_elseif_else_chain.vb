' vybe-test: vb/vb_comprehensive/if_elseif_else_chain
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim x As Integer = 5
        If x > 10 Then
            Console.WriteLine("big")
        ElseIf x > 3 Then
            Console.WriteLine("medium")
        Else
            Console.WriteLine("small")
        End If
    End Sub
End Module
