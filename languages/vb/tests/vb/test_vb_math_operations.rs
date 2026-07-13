use super::helpers::run_vb;

#[test]
fn math_basic_operations() {
    let out = run_vb(
        r#"
Imports System.Math

Module M
    Sub Main()
        Console.WriteLine(Abs(-10))
        Console.WriteLine(Max(5, 10))
        Console.WriteLine(Min(5, 10))
        Console.WriteLine(Pow(2, 3))
        Console.WriteLine(Sqrt(16))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "10", "5", "8", "4"]);
}

#[test]
fn math_rounding_methods() {
    let out = run_vb(
        r#"
Imports System.Math

Module M
    Sub Main()
        Console.WriteLine(Round(2.5)) ' Banker's rounding
        Console.WriteLine(Round(3.5))
        Console.WriteLine(Ceiling(2.1))
        Console.WriteLine(Floor(2.9))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4", "3", "2"]);
}
