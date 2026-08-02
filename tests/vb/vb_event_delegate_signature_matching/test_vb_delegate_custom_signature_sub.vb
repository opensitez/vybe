' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_custom_signature_sub
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Delegate Sub MathOp(a As Integer, b As Integer, ByRef result As Integer)

Module Program
    Private Sub Add(a As Integer, b As Integer, ByRef result As Integer)
        result = a + b
    End Sub

    Sub Main()
        Dim op As MathOp = AddressOf Add
        Dim res As Integer = 0
        op(10, 20, res)
        __Check(CStr(res), "30")
    End Sub
End Module
