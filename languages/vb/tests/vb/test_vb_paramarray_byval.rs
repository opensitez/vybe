use super::helpers::run_vb;

#[test]
fn paramarray_byval() {
    let out = run_vb(
        r#"
Module M
    ' ParamArray is always implicitly ByVal in modern VB, but you can explicitly specify it
    Sub PrintAll(ByVal ParamArray items() As Integer)
        Console.WriteLine(items.Length)
        For Each item In items
            Console.WriteLine(item)
        Next
    End Sub

    Sub Main()
        PrintAll(10, 20, 30)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "10", "20", "30"]);
}
