' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_hierarchy_deep_inheritance_resolution
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

Interface IA : Sub ActA() : End Interface
Interface IB : Inherits IA : Sub ActB() : End Interface
Interface IC : Inherits IB : Sub ActC() : End Interface

Class DeepImpl
    Implements IC
    Public Sub ActA() Implements IA.ActA : __Check(CStr("A"), "A") : End Sub
    Public Sub ActB() Implements IB.ActB : __Check(CStr("B"), "B") : End Sub
    Public Sub ActC() Implements IC.ActC : __Check(CStr("C"), "C") : End Sub
End Class

Module Program
    Sub Main()
        Dim c As IC = New DeepImpl()
        c.ActA() : c.ActB() : c.ActC()
    End Sub
End Module
