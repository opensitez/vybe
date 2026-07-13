use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: GoTo and Labels
// ═══════════════════════════════════════════════════════════

#[test]
fn goto_forward_jump() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Start")
        GoTo SkipThis
        
        Console.WriteLine("Should not print")
        
SkipThis:
        Console.WriteLine("End")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start", "End"]);
}

#[test]
fn goto_backward_loop() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i As Integer = 0
        
LoopStart:
        If i = 3 Then
            GoTo Done
        End If
        
        Console.WriteLine(i)
        i = i + 1
        GoTo LoopStart
        
Done:
        Console.WriteLine("Finished")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "Finished"]);
}

#[test]
fn goto_multiple_labels() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 2
        If x = 1 Then GoTo Label1
        If x = 2 Then GoTo Label2
        If x = 3 Then GoTo Label3
        
Label1:
        Console.WriteLine("L1")
        Exit Sub
Label2:
        Console.WriteLine("L2")
        Exit Sub
Label3:
        Console.WriteLine("L3")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["L2"]);
}
