' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_implementation_in_abstract_class
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IService
    Sub Execute()
End Interface

MustInherit Class BaseService
    Implements IService
    Public MustOverride Sub Execute() Implements IService.Execute
End Class

Class CustomService
    Inherits BaseService
    Public Overrides Sub Execute()
        __Check(CStr("Custom Execution"), "Custom Execution")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IService = New CustomService()
        s.Execute()
    End Sub
End Module
