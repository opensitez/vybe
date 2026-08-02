' vybe-test: vb/vb_istrue_isfalse_operators/istrue_isfalse_operators
' origin: languages/vb/tests/vb/test_vb_istrue_isfalse_operators.rs

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

Class Truthy
    Public Value As Integer
    
    Public Shared Operator IsTrue(ByVal obj As Truthy) As Boolean
        Return obj.Value > 0
    End Operator
    
    Public Shared Operator IsFalse(ByVal obj As Truthy) As Boolean
        Return obj.Value <= 0
    End Operator
End Class

Module M
    Sub Main()
        Dim t1 As New Truthy() With {.Value = 10}
        Dim t2 As New Truthy() With {.Value = -5}
        
        If t1 Then
            __Check(CStr("t1 is true"), "t1 is true")
        End If
        
        If Not t2 Then
            __Check(CStr("t2 is false"), "t2 is false")
        End If
    End Sub
End Module
