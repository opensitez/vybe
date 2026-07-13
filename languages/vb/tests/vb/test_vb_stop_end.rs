use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Stop and End Statements
// ═══════════════════════════════════════════════════════════

#[test]
fn statement_stop() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Before Stop")
        ' Stop suspends execution (in a debugger), but compiler supports parsing it.
        ' Without a debugger attached, behavior varies, but we check parsing/compilation.
        ' Since we don't want to actually suspend our test runner, we will just parse it
        ' but not hit it, or let it compile. Wait, Stop sometimes terminates or throws.
        ' Let's just put it in a non-executed block to verify parser.
        If False Then
            Stop
        End If
        Console.WriteLine("After Stop")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Before Stop", "After Stop"]);
}

#[test]
fn statement_end() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Before End")
        If False Then
            ' End terminates execution immediately
            End
        End If
        Console.WriteLine("After End")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Before End", "After End"]);
}
