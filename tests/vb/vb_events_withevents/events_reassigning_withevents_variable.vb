' vybe-test: vb/vb_events_withevents/events_reassigning_withevents_variable
' origin: languages/vb/tests/vb/test_vb_events_withevents.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Class Worker
    Public ID As Integer
    Public Event Working(id As Integer)
    
    Public Sub New(i As Integer)
        ID = i
    End Sub
    
    Public Sub Work()
        RaiseEvent Working(ID)
    End Sub
End Class

Module M
    Private WithEvents ActiveWorker As Worker
    
    Private Sub OnWorking(id As Integer) Handles ActiveWorker.Working
        __P(CStr("Worker " & id & " is working"))
    End Sub
    
    Sub Main()
        Dim w1 As New Worker(1)
        Dim w2 As New Worker(2)
        
        ActiveWorker = w1
        w1.Work()
        
        ' Reassigning automatically unhooks w1 and hooks w2
        ActiveWorker = w2
        w1.Work() ' Should NOT trigger the handler
        w2.Work() ' SHOULD trigger the handler
        __Check("Worker 1 is working
Worker 2 is working")
    End Sub
End Module
