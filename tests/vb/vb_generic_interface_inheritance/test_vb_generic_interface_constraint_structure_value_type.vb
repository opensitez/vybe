' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_constraint_structure_value_type
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface INumberBox(Of T As Structure)
    Property Value As T
End Interface

Class IntBox
    Implements INumberBox(Of Integer)
    Public Property Value As Integer Implements INumberBox(Of Integer).Value
End Class

Module Program
    Sub Main()
        Dim b As INumberBox(Of Integer) = New IntBox() With {.Value = 99}
        __Check(CStr(b.Value), "99")
    End Sub
End Module
