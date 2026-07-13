use super::helpers::run_vb;

#[test]
fn array_advanced_methods() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {5, 2, 8, 1, 9}
        
        System.Array.Sort(arr)
        Console.WriteLine(arr(0))
        Console.WriteLine(arr(arr.Length - 1))
        
        System.Array.Reverse(arr)
        Console.WriteLine(arr(0))
        
        Dim idx = System.Array.IndexOf(arr, 2)
        Console.WriteLine(idx)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "9", "9", "3"]);
}

#[test]
fn array_copy() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim src() As Integer = {1, 2, 3, 4, 5}
        Dim dest(4) As Integer
        
        System.Array.Copy(src, 1, dest, 2, 2)
        
        For Each v In dest
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "0", "2", "3", "0"]);
}
