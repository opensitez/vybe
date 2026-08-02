' vybe-test: vb/vb_module_alias_imports/module_alias_generic_list_growth_and_index
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

Imports Gen = System.Collections.Generic

Module M
    Sub Main()
        Dim values As New Gen.List(Of Integer)()
        values.Add(2)
        values.Add(4)
        values.Add(6)
        __Check(CStr(values.Count), "3")
        __Check(CStr(values(1)), "4")
    End Sub
End Module
