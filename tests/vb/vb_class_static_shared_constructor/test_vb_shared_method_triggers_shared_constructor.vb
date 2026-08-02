' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_method_triggers_shared_constructor
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

Class Utility
    Public Shared Initialized As Boolean = False
    Shared Sub New()
        Initialized = True
    End Sub
    Public Shared Function Ping() As String
        Return "Pong"
    End Function
End Class

Module Program
    Sub Main()
        Dim res = Utility.Ping()
        __Check(CStr(res & "|Initialized=" & Utility.Initialized), "Pong|Initialized=True")
    End Sub
End Module
