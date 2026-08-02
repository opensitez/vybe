' vybe-test: vb/vb_spec_arrays_collections/array_spec_erase_resets_fixed_integer_array_values
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

Module M : Sub Main() : Dim values(2) As Integer : values(0)=1 : values(1)=2 : values(2)=3 : Erase values : __Check(CStr(values(0)), "0") : __Check(CStr(values(2)), "0") : End Sub : End Module
