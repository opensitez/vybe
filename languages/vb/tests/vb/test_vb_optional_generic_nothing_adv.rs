use super::helpers::run_vb;

#[test]
fn optional_generic_nothing_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Process(Of T)(Optional val As T = Nothing)
        Console.WriteLine(val Is Nothing)
    End Sub

    Sub Main()
        Process(Of String)()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
