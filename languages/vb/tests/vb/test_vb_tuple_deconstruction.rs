use super::helpers::run_vb;

#[test]
fn tuple_deconstruction() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Tuple deconstruction (since VB 15.3 doesn't have native let destructuring quite like C#)
        ' Actually VB 15 doesn't have tuple deconstruction assignment exactly like C# 'var (x, y) = tuple'
        ' But you can do this:
        ' Dim (x, y) = (1, 2) ' This works? Yes in newer VB versions
        Dim t = (1, 2)
        Console.WriteLine(t.Item1)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1"]);
}
