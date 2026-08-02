' vybe-test: vb/vb_spec_arrays_collections/array_spec_lbound_reports_zero_for_array_function_result
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

Module M : Sub Main() : Dim items = Array("a", "b", "c") : __Check(CStr(LBound(items)), "0") : End Sub : End Module
