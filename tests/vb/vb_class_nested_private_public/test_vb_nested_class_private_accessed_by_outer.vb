' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_private_accessed_by_outer
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
    Private Class PrivateInner
        Public Data As String = "SecretData"
    End Class

    Public Function GetInnerData() As String
        Dim inner As New PrivateInner()
        Return inner.Data
    End Function
End Class

Module Program
    Sub Main()
        Dim o As New Outer()
        __Check(CStr(o.GetInnerData()), "SecretData")
    End Sub
End Module
