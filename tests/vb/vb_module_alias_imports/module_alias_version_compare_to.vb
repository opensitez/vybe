' vybe-test: vb/vb_module_alias_imports/module_alias_version_compare_to
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

Imports V = System.Version

Module M
    Sub Main()
        Dim one As New V(1, 2, 3, 4)
        Dim two As New V(1, 2, 4, 4)
        __Check(CStr(CStr(one.CompareTo(two))), "-1")
        __Check(CStr(one.ToString()), "1.2.3.4")
    End Sub
End Module
