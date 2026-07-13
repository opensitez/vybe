use super::helpers::run_vb;

#[test]
fn redim_jagged_array() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr()() As Integer
        
        ReDim arr(1)
        ReDim arr(0)(2)
        ReDim arr(1)(1)
        
        arr(0)(2) = 10
        arr(1)(1) = 20
        
        Console.WriteLine(arr(0)(2))
        Console.WriteLine(arr(1)(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}
