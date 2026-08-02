' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_private_protected_mix
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

Class Base
    Protected Private Class Hidden
        Public ReadOnly Info As String = "HiddenInfo"
    End Class
    Public Function ReadHidden() As String
        Dim h As New Hidden()
        Return h.Info
    End Function
End Class

Module Program
    Sub Main()
        Dim b As New Base()
        __Check(CStr(b.ReadHidden()), "HiddenInfo")
    End Sub
End Module
