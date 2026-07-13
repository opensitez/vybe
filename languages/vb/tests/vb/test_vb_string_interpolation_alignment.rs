use super::helpers::run_vb;

#[test]
fn string_interpolation_alignment() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 42
        
        ' String interpolation with alignment
        Dim s = $"[{val,5}]"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[   42]"]);
}
