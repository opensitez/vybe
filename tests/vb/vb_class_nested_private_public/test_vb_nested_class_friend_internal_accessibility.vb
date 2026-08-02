' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_friend_internal_accessibility
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

Class Parent
    Friend Class InternalHelper
        Public Shared Function Help() As String
            Return "Internal Help"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        __Check(CStr(Parent.InternalHelper.Help()), "Internal Help")
    End Sub
End Module
