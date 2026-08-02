' vybe-test: vb/vb_nullable_types/nullable_value_types_operators
' origin: languages/vb/tests/vb/test_vb_nullable_types.rs

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

Module M
    Sub Main()
        Dim a As Integer? = 5
        Dim b As Integer? = Nothing
        
        ' Lifted operators: If one operand is Nothing, result is Nothing
        Dim result As Integer? = a + b
        
        __Check(CStr(result.HasValue), "False")
        
        ' Coalescing with If operator
        __Check(CStr(If(result, -1)), "-1")
    End Sub
End Module
