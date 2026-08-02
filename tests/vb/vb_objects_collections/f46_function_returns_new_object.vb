' vybe-test: vb/vb_objects_collections/f46_function_returns_new_object
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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

Class Result
    Public Status As String
End Class
Function GetResult() As Result
    Dim r As New Result()
    r.Status = "OK"
    Return r
End Function
Dim res As Result = GetResult()
__Check(CStr(res.Status), "OK")
