use super::helpers::run_vb;

#[test]
fn erase_statement() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim fixedArray(2) As Integer
        Dim dynArray() As Integer = {1, 2, 3}
        
        ' Erase clears the array
        Erase fixedArray ' Reinitializes elements to default (0)
        Erase dynArray   ' Sets the reference to Nothing
        
        Console.WriteLine(fixedArray(0))
        Console.WriteLine(dynArray Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "True"]);
}
