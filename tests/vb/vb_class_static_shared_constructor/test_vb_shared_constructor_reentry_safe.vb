' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_reentry_safe
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class CircularA
    Public Shared Value As Integer
    Shared Sub New()
        Value = CircularB.Value + 10
    End Sub
End Class

Class CircularB
    Public Shared Value As Integer
    Shared Sub New()
        Value = 5
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(CircularA.Value), "15")
    End Sub
End Module
