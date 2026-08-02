' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_initializes_shared_fields
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

Class Config
    Public Shared ReadOnly StartTime As String
    Shared Sub New()
        StartTime = "2025-01-01 00:00:00"
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(Config.StartTime), "2025-01-01 00:00:00")
    End Sub
End Module
