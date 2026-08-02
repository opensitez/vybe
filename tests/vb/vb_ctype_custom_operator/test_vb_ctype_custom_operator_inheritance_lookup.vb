' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_custom_operator_inheritance_lookup
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Class BaseVal
    Public X As Integer
    Public Sub New(val As Integer)
        X = val
    End Sub

    Public Shared Widening Operator CType(v As Integer) As BaseVal
        Return New BaseVal(v)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim bv As BaseVal = CType(50, BaseVal)
        __Check(CStr(bv.X), "50")
    End Sub
End Module
