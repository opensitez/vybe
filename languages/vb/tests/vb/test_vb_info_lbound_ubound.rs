use super::helpers::run_vb;

#[test]
fn info_lbound_ubound() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Array bounds functions
        Dim arr(2) As Integer
        
        Console.WriteLine(LBound(arr)) ' Usually 0
        Console.WriteLine(UBound(arr)) ' 2
        
        ' Multi-dimensional bounds
        Dim matrix(3, 4) As Integer
        Console.WriteLine(UBound(matrix, 1)) ' Rank 1 -> 3
        Console.WriteLine(UBound(matrix, 2)) ' Rank 2 -> 4
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "2", "3", "4"]);
}
