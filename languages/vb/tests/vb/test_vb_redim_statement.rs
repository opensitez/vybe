use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ReDim and ReDim Preserve
// ═══════════════════════════════════════════════════════════

#[test]
fn redim_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer
        ReDim arr(2)
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        Console.WriteLine(arr(1))
        
        ReDim arr(4)
        Console.WriteLine(arr(1)) ' Should be 0 because elements are reset
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20", "0"]);
}

#[test]
fn redim_preserve() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer
        ReDim arr(2)
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        
        ReDim Preserve arr(4)
        arr(3) = 40
        arr(4) = 50
        
        Console.WriteLine(arr(1))
        Console.WriteLine(arr(4))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20", "50"]);
}

#[test]
fn redim_preserve_shrink() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3, 4, 5}
        ReDim Preserve arr(2)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn redim_multidimensional() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim grid(,) As Integer
        ReDim grid(1, 1)
        grid(0, 0) = 1
        grid(1, 1) = 5
        Console.WriteLine(grid(1, 1))
        
        ' ReDim Preserve can only change the last dimension
        ReDim Preserve grid(1, 2)
        grid(1, 2) = 10
        Console.WriteLine(grid(1, 2))
        Console.WriteLine(grid(0, 0)) ' Still 1
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "10", "1"]);
}
