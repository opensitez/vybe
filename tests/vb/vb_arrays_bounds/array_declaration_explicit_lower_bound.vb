' vybe-test: vb/vb_arrays_bounds/array_declaration_explicit_lower_bound
' origin: languages/vb/tests/vb/test_vb_arrays_bounds.rs

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

Module M
    Sub Main()
        ' VB supports explicit bounds 0 To N in declarations
        Dim arr(0 To 4) As Integer
        __Check(CStr(arr.Length), "5")
    End Sub
End Module
