use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Implicit / Explicit Conversion Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn conversion_operators() {
    let out = run_vb(
        r#"
Structure Digit
    Public Value As Byte
    
    Public Sub New(val As Byte)
        Value = val
    End Sub
    
    ' Widening (Implicit) conversion
    Public Shared Widening Operator CType(d As Digit) As Integer
        Return CInt(d.Value)
    End Operator
    
    ' Narrowing (Explicit) conversion
    Public Shared Narrowing Operator CType(i As Integer) As Digit
        Return New Digit(CByte(i Mod 10))
    End Operator
End Structure

Module M
    Sub Main()
        Dim d As New Digit(5)
        
        ' Implicit conversion to Integer
        Dim num As Integer = d
        Console.WriteLine(num)
        
        ' Explicit conversion from Integer to Digit
        Dim d2 As Digit = CType(23, Digit)
        Console.WriteLine(d2.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "3"]);
}
