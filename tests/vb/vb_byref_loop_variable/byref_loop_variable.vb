' vybe-test: vb/vb_byref_loop_variable/byref_loop_variable
' origin: languages/vb/tests/vb/test_vb_byref_loop_variable.rs

Module M
    Sub ModifyByRef(ByRef val As Integer)
        val += 10
    End Sub

    Sub Main()
        ' VB.NET allows passing loop variables ByRef, but modifying it inside the method 
        ' behaves exactly like changing the loop variable directly.
        For i As Integer = 1 To 2
            ModifyByRef(i)
            Console.WriteLine(i)
        Next
    End Sub
End Module
