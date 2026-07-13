use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Type Characters
// ═══════════════════════════════════════════════════════════

#[test]
fn type_characters_variables_and_literals() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Integer type character is %
        Dim num% = 100
        
        ' Long type character is &
        Dim bigNum& = 9999999999
        
        ' Decimal type character is @
        Dim money@ = 99.99@
        
        ' Single type character is !
        Dim float! = 3.14!
        
        ' String type character is $
        Dim text$ = "Hello"
        
        Console.WriteLine(num.GetType().Name)
        Console.WriteLine(bigNum.GetType().Name)
        Console.WriteLine(money.GetType().Name)
        Console.WriteLine(float.GetType().Name)
        Console.WriteLine(text.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Int32", "Int64", "Decimal", "Single", "String"]);
}
