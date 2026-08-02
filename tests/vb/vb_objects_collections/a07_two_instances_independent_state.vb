' vybe-test: vb/vb_objects_collections/a07_two_instances_independent_state
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

Class Widget
    Public Color As String
End Class
Dim w1 As New Widget()
Dim w2 As New Widget()
w1.Color = "red"
w2.Color = "blue"
__Check(CStr(w1.Color), "red")
__Check(CStr(w2.Color), "blue")
