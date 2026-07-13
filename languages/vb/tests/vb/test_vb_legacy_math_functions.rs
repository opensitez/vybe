use super::helpers::run_vb;

#[test]
fn legacy_math_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Initialize random-number generator
        Randomize(42)
        
        ' Rnd returns a Single less than 1 but greater than or equal to 0
        Dim val1 = Rnd()
        Console.WriteLine(val1 >= 0 AndAlso val1 < 1)
        
        ' Int returns the integer portion of a number
        Console.WriteLine(Int(12.34))
        Console.WriteLine(Int(-12.34)) ' Int rounds down (-13)
        
        ' Fix returns the integer portion of a number
        Console.WriteLine(Fix(12.34))
        Console.WriteLine(Fix(-12.34)) ' Fix truncates towards zero (-12)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "12", "-13", "12", "-12"]);
}
