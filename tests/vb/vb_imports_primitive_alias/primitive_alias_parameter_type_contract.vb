' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_parameter_type_contract
' origin: languages/vb/tests/vb/test_vb_imports_primitive_alias.rs

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

Imports Counter = System.Int32

Module M
    Sub Grow(ByVal x As Counter)
        __Check(CStr(x + 1), "42")
    End Sub

    Sub Main()
        Dim value As Counter = 41
        Grow(value)
    End Sub
End Module
