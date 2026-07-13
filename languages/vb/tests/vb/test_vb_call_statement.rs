use super::helpers::run_vb;

#[test]
fn call_statement_basic() {
    let out = run_vb(
        r#"
Module M
    Sub PrintMessage(msg As String)
        Console.WriteLine(msg)
    End Sub

    Function GetValue() As Integer
        Console.WriteLine("Side effect")
        Return 42
    End Function

    Sub Main()
        ' The Call keyword allows calling a Sub or Function, discarding the return value if any
        Call PrintMessage("Hello using Call")
        Call GetValue()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello using Call", "Side effect"]);
}

#[test]
fn call_statement_with_object() {
    let out = run_vb(
        r#"
Class Logger
    Public Sub Log(msg As String)
        Console.WriteLine("Log: " & msg)
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        Call l.Log("Test Call")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Log: Test Call"]);
}
