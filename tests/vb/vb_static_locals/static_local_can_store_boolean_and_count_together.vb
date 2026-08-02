' vybe-test: vb/vb_static_locals/static_local_can_store_boolean_and_count_together
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

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
    Function Snapshot() As String
        Static flag As Boolean = False
        Static count As Integer = 0
        flag = Not flag
        If flag Then
            count = count + 1
        End If
        Return flag & ":" & count
    End Function

    Sub Main()
        __Check(CStr(Snapshot()), "true:1")
        __Check(CStr(Snapshot()), "false:1")
        __Check(CStr(Snapshot()), "true:2")
    End Sub
End Module
