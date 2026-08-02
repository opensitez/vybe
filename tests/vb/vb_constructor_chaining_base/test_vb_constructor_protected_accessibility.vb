' vybe-test: vb/vb_constructor_chaining_base/test_vb_constructor_protected_accessibility
' origin: languages/vb/tests/vb/test_vb_constructor_chaining_base.rs

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

Class ProtectedBase
    Protected Sub New()
        __Check(CStr("Protected Ctor Called"), "Protected Ctor Called")
    End Sub
End Class

Class PublicDerived
    Inherits ProtectedBase
    Public Sub New()
        MyBase.New()
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New PublicDerived()
    End Sub
End Module
