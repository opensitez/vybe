' vybe-test: vb/vb_spec_arrays_collections/array_spec_redim_preserve_can_be_used_multiple_times
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

Module M : Sub Main() : Dim items() As Integer = {1} : ReDim Preserve items(1) : ReDim Preserve items(2) : items(2)=5 : __Check(CStr(items(0)), "1") : __Check(CStr(items(2)), "5") : End Sub : End Module
