use super::helpers::run_vb;

#[test]
fn redim_multidimensional_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr(,) As Integer
        
        ' ReDim changes bounds
        ReDim arr(2, 2)
        arr(1, 1) = 5
        
        ' Preserve keeps data, but only the last dimension can change
        ReDim Preserve arr(2, 4)
        Console.WriteLine(arr(1, 1))
        Console.WriteLine(arr.GetLength(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "5"]);
}
