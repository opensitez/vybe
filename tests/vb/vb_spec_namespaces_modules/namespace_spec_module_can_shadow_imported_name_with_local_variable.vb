' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_module_can_shadow_imported_name_with_local_variable
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

Namespace Demo
    Public Module Util
        Public Function Value() As String
            Return "module"
        End Function
    End Module
End Namespace
Imports Demo
Module M
    Sub Main()
        Dim Util As String = "local"
        __Check(CStr(Util), "local")
    End Sub
End Module
