use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Events — Event/RaiseEvent, AddHandler
// ═══════════════════════════════════════════════════════════

#[test]
#[ignore]
fn event_raiseevent() {
    let out = run_vb(r#"
Class Timer
    Public Event Tick()
    Public Sub Fire()
        RaiseEvent Tick()
    End Sub
End Class

Module M
    Sub OnTick()
        Console.WriteLine("tick!")
    End Sub
    Sub Main()
        Dim t As New Timer()
        AddHandler t.Tick, AddressOf OnTick
        t.Fire()
        t.Fire()
    End Sub
End Module
"#);
    assert_eq!(out, vec!["tick!", "tick!"]);
}

#[test]
#[ignore]
fn addhandler_removehandler() {
    let out = run_vb(r#"
Class Button
    Public Event Click()
    Public Sub DoClick()
        RaiseEvent Click()
    End Sub
End Class

Module M
    Sub OnClick()
        Console.WriteLine("clicked")
    End Sub
    Sub Main()
        Dim btn As New Button()
        AddHandler btn.Click, AddressOf OnClick
        btn.DoClick()
        RemoveHandler btn.Click, AddressOf OnClick
        btn.DoClick()
        Console.WriteLine("done")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["clicked", "done"]);
}

#[test]
#[ignore]
fn multiple_handlers() {
    let out = run_vb(r#"
Class Notifier
    Public Event Notify()
    Public Sub Fire()
        RaiseEvent Notify()
    End Sub
End Class

Module M
    Sub Handler1()
        Console.WriteLine("h1")
    End Sub
    Sub Handler2()
        Console.WriteLine("h2")
    End Sub
    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Notify, AddressOf Handler1
        AddHandler n.Notify, AddressOf Handler2
        n.Fire()
    End Sub
End Module
"#);
    assert_eq!(out, vec!["h1", "h2"]);
}
