' vybe-test: vb/vb_basic/if_elseif_else
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
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
