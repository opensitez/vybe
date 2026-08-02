' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_namespace_type_can_use_module_helper_in_same_namespace
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
        Public Function GetName() As String
            Return "demo"
        End Function
    End Module
    Public Class Box
        Public Function Name() As String
            Return Util.GetName()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        __Check(CStr((New Demo.Box()).Name()), "demo")
    End Sub
End Module
