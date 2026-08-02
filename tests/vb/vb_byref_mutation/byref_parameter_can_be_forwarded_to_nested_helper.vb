' vybe-test: vb/vb_byref_mutation/byref_parameter_can_be_forwarded_to_nested_helper
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
    Sub Inner(ByRef value As Integer)
        value = value * 2
    End Sub

    Sub Outer(ByRef value As Integer)
        Inner(value)
    End Sub

    Sub Main()
        Dim value As Integer = 6
        Outer(value)
        __Check(CStr(value), "12")
    End Sub
End Module
