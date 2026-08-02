' vybe-test: vb/vb_byref_mutation/byref_integer_parameter_can_assign_absolute_value
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
    Sub ResetValue(ByRef value As Integer)
        value = 42
    End Sub

    Sub Main()
        Dim x As Integer = 3
        ResetValue(x)
        __Check(CStr(x), "42")
    End Sub
End Module
