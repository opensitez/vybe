' vybe-test: vb/vb_static_locals/static_local_can_be_used_from_instance_method
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
    Public Function NextId() As Integer
        Static current As Integer = 100
        current = current + 1
        Return current
    End Function
End Class

Module M
    Sub Main()
        Dim worker As New Worker()
        __Check(CStr(worker.NextId()), "101")
        __Check(CStr(worker.NextId()), "102")
    End Sub
End Module
