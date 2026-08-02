' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_can_be_received_from_function_into_local
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

Module M : Function Build() As String() : Return New String() {"x","y"} : End Function : Sub Main() : Dim items() As String = Build() : __Check(CStr(items(0)), "x") : End Sub : End Module
