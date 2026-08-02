' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_accessing_outer_private_static_member
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class Outer
    Private Shared OuterSecret As String = "TopSecret"

    Public Class Inner
        Public Function ReadSecret() As String
            Return OuterSecret
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim inObj As New Outer.Inner()
        __Check(CStr(inObj.ReadSecret()), "TopSecret")
    End Sub
End Module
