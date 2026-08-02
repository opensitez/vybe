' vybe-test: vb/vb_module_alias_imports/module_alias_array_clone_preserved
' origin: languages/vb/tests/vb/test_vb_module_alias_imports.rs

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

Imports Arr = System

Module M
    Sub Main()
        Dim source() As Integer = {1, 2, 3}
        Dim copy() As Integer = Arr.ConvertAll(source, Function(x) x + 2)
        copy(1) = 9
        __Check(CStr(source(1)), "2")
        __Check(CStr(copy(1)), "9")
    End Sub
End Module
