' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_of_dictionaries_can_store_lookup_objects
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

Module M : Sub Main() : Dim maps() As Dictionary(Of String, Integer) = { New Dictionary(Of String, Integer)() } : maps(0).Add("x", 7) : __Check(CStr(maps(0).Item("x")), "7") : End Sub : End Module
