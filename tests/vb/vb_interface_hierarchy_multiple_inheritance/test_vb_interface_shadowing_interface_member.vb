' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_shadowing_interface_member
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

Interface IBase
    Sub Render()
End Interface

Interface IDerived
    Inherits IBase
    Shadows Sub Render()
End Interface

Class Window
    Implements IDerived
    Public Sub BaseRender() Implements IBase.Render
        __Check(CStr("IBase Render"), "IBase Render")
    End Sub
    Public Sub DerivedRender() Implements IDerived.Render
        __Check(CStr("IDerived Render"), "IDerived Render")
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New Window()
        Dim b As IBase = w
        Dim d As IDerived = w
        b.Render()
        d.Render()
    End Sub
End Module
