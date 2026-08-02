' vybe-test: vb/vb_interop/a09_overridden_method_replaces_parent
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Base
    Public Function Name() As String
        Return "Base"
    End Function
End Class
Public Class Derived
    Inherits Base
    Public Overrides Function Name() As String
        Return "Derived"
    End Function
End Class
Dim d As New Derived()
__Check(CStr(d.Name()), "Derived")
