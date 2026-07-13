use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: WithEvents and Handles Keyword
// ═══════════════════════════════════════════════════════════

#[test]
fn events_withevents_handles() {
    let out = run_vb(
        r#"
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
        Console.WriteLine("Completed: " & count.ToString())
    End Sub
    
    Sub Main()
        Processor = New DataProcessor()
        Processor.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Completed: 42"]);
}

#[test]
fn events_handles_multiple_events() {
    let out = run_vb(
        r#"
Class Button
    Public Event Click()
    Public Event MouseEnter()
    
    Public Sub SimulateClick()
        RaiseEvent Click()
    End Sub
    
    Public Sub SimulateHover()
        RaiseEvent MouseEnter()
    End Sub
End Class

Module M
    Private WithEvents Btn As Button
    
    ' One handler method handling multiple events
    Private Sub OnUserInteraction() Handles Btn.Click, Btn.MouseEnter
        Console.WriteLine("Interaction occurred")
    End Sub
    
    Sub Main()
        Btn = New Button()
        Btn.SimulateClick()
        Btn.SimulateHover()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Interaction occurred", "Interaction occurred"]);
}

#[test]
fn events_reassigning_withevents_variable() {
    let out = run_vb(
        r#"
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
        Console.WriteLine("Worker " & id & " is working")
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
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Worker 1 is working", "Worker 2 is working"]);
}
