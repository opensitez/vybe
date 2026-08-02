' vybe-test: vb/vb_events_withevents/events_withevents_handles
' origin: languages/vb/tests/vb/test_vb_events_withevents.rs

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

Class DataProcessor
    Public Event ProcessCompleted(count As Integer)
    
    Public Sub DoWork()
        RaiseEvent ProcessCompleted(42)
    End Sub
End Class

Module M
    ' WithEvents declares an object variable that responds to events
    Private WithEvents Processor As DataProcessor
    
    ' Handles links the event to this specific method
    Private Sub OnProcessCompleted(count As Integer) Handles Processor.ProcessCompleted
        __Check(CStr("Completed: " & count.ToString()), "Completed: 42")
    End Sub
    
    Sub Main()
        Processor = New DataProcessor()
        Processor.DoWork()
    End Sub
End Module
