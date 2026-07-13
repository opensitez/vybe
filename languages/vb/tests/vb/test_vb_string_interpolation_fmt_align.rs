use super::helpers::run_vb;

#[test]
fn string_interpolation_fmt_align() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val As Double = 42.5
        
        ' String interpolation with alignment and format specifier
        Dim s = $"[{val,10:F2}]"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[     42.50]"]);
}
