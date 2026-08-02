' vybe-test: vb/vb_operator_istrue_isfalse/operator_istrue_isfalse
' origin: languages/vb/tests/vb/test_vb_operator_istrue_isfalse.rs

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

Class TriState
    Public Value As Integer ' 0 = False, 1 = True, -1 = Unknown
    
    Public Shared Operator IsTrue(t As TriState) As Boolean
        Return t.Value = 1
    End Operator
    
    Public Shared Operator IsFalse(t As TriState) As Boolean
        Return t.Value = 0
    End Operator
End Class

Module M
    Sub Main()
        Dim t As New TriState() With {.Value = 1}
        
        ' Relies on IsTrue operator
        If t Then
            __Check(CStr("True"), "True")
        End If
    End Sub
End Module
