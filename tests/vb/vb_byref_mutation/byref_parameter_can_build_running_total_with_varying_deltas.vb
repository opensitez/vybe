' vybe-test: vb/vb_byref_mutation/byref_parameter_can_build_running_total_with_varying_deltas
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
    Sub Add(ByRef total As Integer, amount As Integer)
        total = total + amount
    End Sub

    Sub Main()
        Dim total As Integer = 0
        Add(total, 3)
        Add(total, 7)
        Add(total, -2)
        __Check(CStr(total), "8")
    End Sub
End Module
