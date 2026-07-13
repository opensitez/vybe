use super::helpers::run_vb;

#[test]
fn static_in_generic() {
    let out = run_vb(
        r#"
Module M
    Function GetCounter(Of T)() As Integer
        ' Static variables inside generic methods are scoped per generic type parameter
        Static c As Integer = 0
        c += 1
        Return c
    End Function

    Sub Main()
        Console.WriteLine(GetCounter(Of Integer)())
        Console.WriteLine(GetCounter(Of Integer)())
        Console.WriteLine(GetCounter(Of String)())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "1"]);
}
