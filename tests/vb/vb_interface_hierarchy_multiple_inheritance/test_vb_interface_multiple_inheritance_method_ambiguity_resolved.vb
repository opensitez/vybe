' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_multiple_inheritance_method_ambiguity_resolved
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

Interface IX
    Sub Print()
End Interface

Interface IY
    Sub Print()
End Interface

Class Implementation
    Implements IX, IY
    Private Sub PrintX() Implements IX.Print
        __Check(CStr("IX Print"), "IX Print")
    End Sub
    Private Sub PrintY() Implements IY.Print
        __Check(CStr("IY Print"), "IY Print")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New Implementation()
        Dim x As IX = obj
        Dim y As IY = obj
        x.Print()
        y.Print()
    End Sub
End Module
