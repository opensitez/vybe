use super::helpers::run_vb;

#[test]
fn iterator_function_returns_continuation() {
    let out = run_vb(
        r#"
Module Program
    Function Count()
        Yield 1
        Yield 2
    End Function

    Sub Main()
        Console.WriteLine(Count())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn iterator_function_body_stays_lazy() {
    let out = run_vb(
        r#"
Module Program
    Function Loud()
        Console.WriteLine("bad")
        Yield 1
    End Function

    Sub Main()
        Dim g = Loud()
        Console.WriteLine("ok")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ok"]);
}
