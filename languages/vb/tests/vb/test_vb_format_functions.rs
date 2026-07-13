use super::helpers::run_vb;

#[test]
fn format_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Format functions
        Console.WriteLine(Format(12.34, "0.0"))
        
        ' In some locales currency symbol changes, so we just check it doesn't throw
        Dim currencyStr = FormatCurrency(12.34)
        Console.WriteLine(currencyStr.Length > 0)
        
        Dim numStr = FormatNumber(12.34, 1)
        Console.WriteLine(numStr.Length > 0)
        
        Dim pctStr = FormatPercent(0.123, 1)
        Console.WriteLine(pctStr.Length > 0)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12.3", "True", "True", "True"]);
}
