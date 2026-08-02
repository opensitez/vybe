' vybe-test: vb/vb_oop_edges/custom_operator_like
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Class Pattern
    Public Shared Operator Like(obj As Pattern, pattern As String) As Boolean
        Return True
    End Operator
End Class

Module M
    Sub Main()
        Dim p As New Pattern()
        If p Like "*test*" Then
            __Check(CStr("Matched"), "Matched")
        End If
    End Sub
End Module
