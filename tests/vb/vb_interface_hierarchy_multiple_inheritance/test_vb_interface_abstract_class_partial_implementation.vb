' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_abstract_class_partial_implementation
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IFullService
    Sub ActionA()
    Sub ActionB()
End Interface

MustInherit Class PartialService
    Implements IFullService
    Public Sub ActionA() Implements IFullService.ActionA
        __Check(CStr("ActionA Completed"), "ActionA Completed")
    End Sub
    Public MustOverride Sub ActionB() Implements IFullService.ActionB
End Class

Class CompleteService
    Inherits PartialService
    Public Overrides Sub ActionB()
        __Check(CStr("ActionB Completed"), "ActionB Completed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IFullService = New CompleteService()
        s.ActionA()
        s.ActionB()
    End Sub
End Module
