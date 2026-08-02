' vybe-test: vb/vb_objects_collections/a04_object_passed_to_function_modified
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

Class Counter
    Public Value As Integer
End Class
Sub Increment(c As Counter)
    c.Value = c.Value + 1
End Sub
Dim c As New Counter()
c.Value = 10
Increment(c)
__Check(CStr(c.Value), "11")
