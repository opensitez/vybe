' vybe-test: vb/vb_module_alias_imports/module_alias_generic_dictionary_contains_key
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

Imports Ctx = System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Ctx.Dictionary(Of String, Integer)()
        map("one") = 1
        map("two") = 2
        __Check(CStr(map.ContainsKey("two")), "True")
        __Check(CStr(map("two")), "2")
    End Sub
End Module
