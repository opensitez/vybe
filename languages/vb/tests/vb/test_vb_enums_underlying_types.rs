use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Enums (Underlying Types)
// ═══════════════════════════════════════════════════════════

#[test]
fn enum_underlying_type_byte() {
    let out = run_vb(
        r#"
Enum SmallNumber As Byte
    Zero = 0
    One = 1
    Two = 2
End Enum

Module M
    Sub Main()
        Dim s As SmallNumber = SmallNumber.Two
        ' Prints the numeric value when cast to Byte
        Console.WriteLine(CByte(s))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_underlying_type_long() {
    let out = run_vb(
        r#"
Enum BigEnum As Long
    Max = 9223372036854775807
    Min = -9223372036854775808
End Enum

Module M
    Sub Main()
        Console.WriteLine(CLng(BigEnum.Max))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9223372036854775807"]);
}
