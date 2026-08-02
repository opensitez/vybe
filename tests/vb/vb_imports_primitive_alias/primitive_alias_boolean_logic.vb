' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_boolean_logic
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

Imports MyFlag = System.Boolean

Module M
    Sub Main()
        Dim enabled As MyFlag = True
        __Check(CStr(Not enabled), "False")
        __Check(CStr(enabled AndAlso False), "False")
    End Sub
End Module
