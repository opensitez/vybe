' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_explicit_name_aliasing
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Class FileLogger
    Implements ILogger
    Public Sub RecordMessage(msg As String) Implements ILogger.Log
        __Check(CStr("LOG: " & msg), "LOG: System Started")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As ILogger = New FileLogger()
        l.Log("System Started")
    End Sub
End Module
