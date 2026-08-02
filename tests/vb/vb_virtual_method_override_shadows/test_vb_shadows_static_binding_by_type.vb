' vybe-test: vb/vb_virtual_method_override_shadows/test_vb_shadows_static_binding_by_type
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

Class Parent
    Public Sub Show()
        __Check(CStr("Parent Show"), "Child Show")
    End Sub
End Class

Class Child
    Inherits Parent
    Public Shadows Sub Show()
        __Check(CStr("Child Show"), "Parent Show")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Child()
        Dim p As Parent = c
        c.Show()
        p.Show()
    End Sub
End Module
