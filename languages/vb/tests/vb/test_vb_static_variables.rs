use super::helpers::run_vb;

#[test]
fn static_variables() {
    let out = run_vb(
        r#"
Module M
    Function GetNextId() As Integer
        ' Static local variable retains its value between calls
        Static id As Integer = 0
        id += 1
        Return id
    End Function

    Sub Main()
        Console.WriteLine(GetNextId())
        Console.WriteLine(GetNextId())
        Console.WriteLine(GetNextId())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
