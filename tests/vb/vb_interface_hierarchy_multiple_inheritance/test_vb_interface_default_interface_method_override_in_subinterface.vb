' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_default_interface_method_override_in_subinterface
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface ILogger
    Sub Log(msg As String)
End Interface

Interface IAdvancedLogger
    Inherits ILogger
    Sub Log(msg As String, severity As Integer)
End Interface

Class CustomLogger
    Implements IAdvancedLogger
    Public Sub Log(msg As String) Implements ILogger.Log
        __Check(CStr("Basic: " & msg), "Basic: System Start")
    End Sub
    Public Sub Log(msg As String, severity As Integer) Implements IAdvancedLogger.Log
        __Check(CStr("Advanced [" & severity & "]: " & msg), "Advanced [5]: Critical Failure")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As IAdvancedLogger = New CustomLogger()
        l.Log("System Start")
        l.Log("Critical Failure", 5)
    End Sub
End Module
