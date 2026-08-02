' vybe-test: vb/vb_objects_collections/a09_object_field_set_to_another_object
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

Class Inner
    Public Val As Integer
End Class
Class Outer
    Public Child As Inner
End Class
Dim i As New Inner()
i.Val = 99
Dim o As New Outer()
o.Child = i
__Check(CStr(o.Child.Val), "99")
