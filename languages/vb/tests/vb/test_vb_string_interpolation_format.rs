use super::helpers::run_vb;

#[test]
fn string_interpolation_format() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val As Double = 12.3456
        
        ' String interpolation with format specifier
        Dim s = $"Value: {val:F2}"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Value: 12.35"]);
}
