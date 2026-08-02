' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_module_method_can_return_array
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

Module Data
    Public Function Build() As Integer()
        Return New Integer() {1, 2, 3}
    End Function
End Module
Module M
    Sub Main()
        __Check(CStr(Data.Build()(1)), "2")
    End Sub
End Module
