' vybe-test: vb/vb_basic/if_else
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
    Sub Main()
        Dim x As Integer = 3
        If x > 5 Then
            Console.WriteLine("big")
        Else
            Console.WriteLine("small")
        End If
    End Sub
End Module
