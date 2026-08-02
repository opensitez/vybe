' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_multiple_interface_inheritance_conflict_resolution
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Interface ILoggerA
    Sub Log(msg As String)
End Interface

Interface ILoggerB
    Sub Log(msg As String)
End Interface

Class DualLogger
    Implements ILoggerA, ILoggerB

    Private Sub LogA(msg As String) Implements ILoggerA.Log
        __Check(CStr("LoggerA: " & msg), "LoggerA: Message")
    End Sub

    Private Sub LogB(msg As String) Implements ILoggerB.Log
        __Check(CStr("LoggerB: " & msg), "LoggerB: Message")
    End Sub
End Class

Module Program
    Sub Main()
        Dim dl As New DualLogger()
        Dim a As ILoggerA = dl
        Dim b As ILoggerB = dl
        a.Log("Message")
        b.Log("Message")
    End Sub
End Module
