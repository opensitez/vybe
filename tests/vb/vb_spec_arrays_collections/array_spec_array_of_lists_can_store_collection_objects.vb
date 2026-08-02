' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_of_lists_can_store_collection_objects
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

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

Module M : Sub Main() : Dim bags() As List(Of Integer) = { New List(Of Integer)(), New List(Of Integer)() } : bags(1).Add(9) : __Check(CStr(bags(1).Count), "1") : End Sub : End Module
