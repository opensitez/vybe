' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_property_triggers_shared_constructor
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

Class SystemInfo
    Private Shared _version As String
    Shared Sub New()
        _version = "v1.0.0"
    End Sub
    Public Shared ReadOnly Property Version As String
        Get
            Return _version
        End Get
    End Property
End Class

Module Program
    Sub Main()
        __Check(CStr(SystemInfo.Version), "v1.0.0")
    End Sub
End Module
