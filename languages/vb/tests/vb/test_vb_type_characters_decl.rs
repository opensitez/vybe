use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Type Characters in Variable Declarations
// ═══════════════════════════════════════════════════════════

#[test]
fn type_characters_declarations() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Type characters specify the variable's type without an As clause
        Dim a$ = "Hello" ' String
        Dim b% = 42      ' Integer
        Dim c& = 100000L ' Long
        Dim d! = 1.5F    ' Single
        Dim e# = 2.5     ' Double
        Dim f@ = 3.5D    ' Decimal
        
        Console.WriteLine(a.GetType().Name)
        Console.WriteLine(b.GetType().Name)
        Console.WriteLine(c.GetType().Name)
        Console.WriteLine(d.GetType().Name)
        Console.WriteLine(e.GetType().Name)
        Console.WriteLine(f.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["String", "Int32", "Int64", "Single", "Double", "Decimal"]
    );
}
