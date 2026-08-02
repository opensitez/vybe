' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_of_structures_preserves_member_reads
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

Structure Point : Public X As Integer : End Structure : Module M : Sub Main() : Dim points() As Point = { New Point With {.X = 2}, New Point With {.X = 7} } : __Check(CStr(points(1).X), "7") : End Sub : End Module
