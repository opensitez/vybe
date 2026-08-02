' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_of_objects_preserves_field_access
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

Class C : Public Name As String : Public Sub New(v As String) : Name=v : End Sub : End Class : Module M : Sub Main() : Dim items() As C = {New C("a"), New C("b")} : __Check(CStr(items(1).Name), "b") : End Sub : End Module
