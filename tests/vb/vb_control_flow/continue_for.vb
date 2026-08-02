' vybe-test: vb/vb_control_flow/continue_for
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim sum As Integer = 0
        For i As Integer = 1 To 10
            If i Mod 2 <> 0 Then Continue For
            sum = sum + i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
