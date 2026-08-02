' vybe-test: vb/vb_static_locals/static_local_can_be_used_from_shared_method
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

Class Worker
    Public Shared Function NextBatch() As Integer
        Static batch As Integer = 1
        batch = batch * 2
        Return batch
    End Function
End Class

Module M
    Sub Main()
        __Check(CStr(Worker.NextBatch()), "2")
        __Check(CStr(Worker.NextBatch()), "4")
        __Check(CStr(Worker.NextBatch()), "8")
    End Sub
End Module
