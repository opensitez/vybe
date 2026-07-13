use super::helpers::run_vb;

#[test]
fn paramarray_modifier() {
    let out = run_vb(
        r#"
Module M
    ' ParamArray allows passing a variable number of arguments
    Function Sum(ParamArray nums() As Integer) As Integer
        Dim total = 0
        For Each n In nums
            total += n
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Sum())
        Console.WriteLine(Sum(1, 2, 3))
        
        Dim arr() As Integer = {4, 5, 6}
        Console.WriteLine(Sum(arr))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "6", "15"]);
}
