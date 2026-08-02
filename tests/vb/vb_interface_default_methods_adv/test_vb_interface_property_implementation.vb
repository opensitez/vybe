' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_property_implementation
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

Interface INamed
    Property Name As String
End Interface

Class User
    Implements INamed
    Public Property Name As String Implements INamed.Name
End Class

Module Program
    Sub Main()
        Dim n As INamed = New User() With {.Name = "Alice"}
        __Check(CStr(n.Name), "Alice")
    End Sub
End Module
