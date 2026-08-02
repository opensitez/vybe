' vybe-test: vb/vb_casts/trycast_result_can_flow_through_isnothing_check
' origin: languages/vb/tests/vb/test_vb_casts.rs

Module M
    Sub Main()
        Dim boxed As Object = "vb"
        Dim value As String = TryCast(boxed, String)
        If IsNothing(value) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine(value)
        End If
    End Sub
End Module
