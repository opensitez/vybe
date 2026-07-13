use super::helpers::run_vb;

#[test]
fn array_bounds_upper() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' In VB.NET, you specify the upper bound, not the length
        Dim arr(2) As Integer
        
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        
        Console.WriteLine(arr.Length) ' Length is 3
        Console.WriteLine(arr.GetUpperBound(0)) ' Upper bound is 2
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}
