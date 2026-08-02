' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_nested_delegate_type_in_interface
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Interface IProcessor
    Delegate Sub ResultHandler(success As Boolean, payload As String)
    Sub Process(callback As ResultHandler)
End Interface

Class ConcreteProcessor
    Implements IProcessor
    Public Sub Process(callback As IProcessor.ResultHandler) Implements IProcessor.Process
        callback(True, "OK")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IProcessor = New ConcreteProcessor()
        p.Process(Sub(s, msg) __Check(CStr(s & ":" & msg), "True:OK"))
    End Sub
End Module
