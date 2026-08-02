' vybe-test: vb/vb_byref_mutation/byref_parameter_updates_caller_when_used_with_if_branch
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
    Sub Adjust(ByRef value As Integer, shouldBoost As Boolean)
        If shouldBoost Then
            value = value + 10
        Else
            value = value + 1
        End If
    End Sub

    Sub Main()
        Dim x As Integer = 5
        Adjust(x, True)
        Adjust(x, False)
        __Check(CStr(x), "16")
    End Sub
End Module
