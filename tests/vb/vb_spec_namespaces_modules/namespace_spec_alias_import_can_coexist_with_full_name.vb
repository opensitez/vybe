' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_alias_import_can_coexist_with_full_name
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

Namespace Demo.Core
    Public Class Box
        Public Function Name() As String
            Return "box"
        End Function
    End Class
End Namespace
Imports CoreAlias = Demo.Core
Module M
    Sub Main()
        __Check(CStr((New CoreAlias.Box()).Name()), "box")
        __Check(CStr((New Demo.Core.Box()).Name()), "box")
    End Sub
End Module
