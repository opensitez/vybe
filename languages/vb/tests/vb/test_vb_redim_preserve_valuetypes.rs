use super::helpers::run_vb;

#[test]
fn redim_preserve_value_types() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3}
        
        ' ReDim Preserve keeps existing values
        ReDim Preserve arr(4)
        
        arr(3) = 4
        arr(4) = 5
        
        For i As Integer = 0 To arr.Length - 1
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4", "5"]);
}
