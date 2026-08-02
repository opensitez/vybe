' vybe-test: vb/vb_virtual_method_override_shadows/test_vb_mybase_call_virtual_method
' origin: languages/vb/tests/vb/test_vb_virtual_method_override_shadows.rs

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

Class BaseService
    Public Overridable Sub Execute()
        __Check(CStr("Base Service"), "Base Service")
    End Sub
End Class

Class ExtendedService
    Inherits BaseService
    Public Overrides Sub Execute()
        MyBase.Execute()
        __Check(CStr("Extended Service"), "Extended Service")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As BaseService = New ExtendedService()
        s.Execute()
    End Sub
End Module
