' vybe-test: vb/vb_static_locals/static_local_multiple_values_can_coexist_in_one_function
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
        Static count As Integer = 0
        Static text As String = "seed"
        count = count + 1
        text = text & count
        Return count & ":" & text
    End Function

    Sub Main()
        __Check(CStr(Snapshot()), "1:seed1")
        __Check(CStr(Snapshot()), "2:seed12")
    End Sub
End Module
