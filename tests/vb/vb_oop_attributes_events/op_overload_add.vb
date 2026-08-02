' vybe-test: vb/vb_oop_attributes_events/op_overload_add
' origin: languages/vb/tests/vb/test_vb_oop_attributes_events.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Class C: Public V As Integer: Public Shared Operator +(a As C, b As C) As C: Return New C With {.V = a.V + b.V}: End Operator: End Class: Module M: Sub Main(): Dim c1 As New C With {.V = 1}: Dim c2 As New C With {.V = 2}: __Check(CStr((c1 + c2).V), "3"): End Sub: End Module
