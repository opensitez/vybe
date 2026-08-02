' vybe-test: vb/vb_widening_operator_custom/widening_operator_custom
' origin: languages/vb/tests/vb/test_vb_widening_operator_custom.rs

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

Class BigNum
    Public Value As Integer
    
    ' Implicit conversion from Integer
    Public Shared Widening Operator CType(i As Integer) As BigNum
        Return New BigNum() With {.Value = i}
    End Operator
End Class

Module M
    Sub Main()
        ' Implicitly triggers Widening Operator
        Dim b As BigNum = 42
        __Check(CStr(b.Value), "42")
    End Sub
End Module
