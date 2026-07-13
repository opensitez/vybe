use super::helpers::run_vb;

#[test]
fn return_array() {
    let out = run_vb(
        r#"
Module M
    ' Method returning an array
    Function GetNames() As String()
        Return {"Alice", "Bob"}
    End Function

    Sub Main()
        Dim names = GetNames()
        Console.WriteLine(names(0))
        Console.WriteLine(names(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "Bob"]);
}
