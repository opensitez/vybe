use super::helpers::run_vb;

#[test]
fn array_lower_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' VB supports specifying the lower bound, though it must be 0 in .NET
        Dim arr(0 To 2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3
        
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}
