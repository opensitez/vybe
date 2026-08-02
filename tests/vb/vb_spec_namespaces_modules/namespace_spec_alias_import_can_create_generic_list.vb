' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_alias_import_can_create_generic_list
' origin: languages/vb/tests/vb/test_vb_spec_namespaces_modules.rs

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

Imports IntList = System.Collections.Generic.List(Of Integer)
Module M
    Sub Main()
        Dim items As New IntList()
        items.Add(3)
        __Check(CStr(items(0)), "3")
    End Sub
End Module
