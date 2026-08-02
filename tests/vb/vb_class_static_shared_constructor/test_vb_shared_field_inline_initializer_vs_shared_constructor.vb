' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_field_inline_initializer_vs_shared_constructor
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

Class OrderOfExecution
    Public Shared Value As Integer = 10
    Shared Sub New()
        Value = 20
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(OrderOfExecution.Value), "20")
    End Sub
End Module
