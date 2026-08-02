' vybe-test: vb/vb_byref_mutation/byref_nested_dispatch_chain_uses_same_alias_in_multiple_methods
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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
    Sub Multiply(ByRef value As Integer, multiplier As Integer)
        value *= multiplier
    End Sub

    Sub Apply(ByRef value As Integer)
        Multiply(value, 2)
        Multiply(value, 3)
    End Sub

    Sub Main()
        Dim value As Integer = 4
        Apply(value)
        __Check(CStr(value), "24")
    End Sub
End Module
