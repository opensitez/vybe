' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_module_property_can_be_accessed_with_namespace_qualification
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
    Public Module Config
        Public Property Name As String
    End Module
End Namespace
Module M
    Sub Main()
        Demo.Config.Name = "cfg"
        __Check(CStr(Demo.Config.Name), "cfg")
    End Sub
End Module
