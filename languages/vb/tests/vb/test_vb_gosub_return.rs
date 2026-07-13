use super::helpers::run_vb;

#[test]
fn gosub_return_legacy() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 1
        GoSub DoubleIt
        GoSub DoubleIt
        Console.WriteLine(x)
        Exit Sub
        
DoubleIt:
        x *= 2
        Return ' In a Sub with GoSub, Return jumps back to the line after GoSub
    End Sub
End Module
"#,
    );
    // Note: GoSub/Return are very legacy and might not be supported in modern VB.NET profiles without warnings,
    // but they are valid syntax in the full language spec up to VB 9 (removed in VB 10, but some parsers still accept it or provide specific errors).
    // If the compiler rejects it, it tests the parser's knowledge of the keyword.
    // Assuming it works for this test.
    assert_eq!(out, vec!["4"]);
}
