' vybe-test: vb/vb_byref_mutation/byref_parameter_mutation_inside_loop_persists_to_caller
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

Module M
    Sub AddRange(ByRef value As Integer)
        For i As Integer = 1 To 4
            value = value + i
        Next
    End Sub

    Sub Main()
        Dim total As Integer = 0
        AddRange(total)
        Console.WriteLine(total)
    End Sub
End Module
