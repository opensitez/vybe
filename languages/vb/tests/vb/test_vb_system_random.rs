use super::helpers::run_vb;

#[test]
fn system_random_next() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random(42) ' Seeded for determinism
        Dim val1 = r.Next(1, 100)
        Dim val2 = r.Next(1, 100)
        
        Console.WriteLine(val1 >= 1 AndAlso val1 < 100)
        Console.WriteLine(val2 >= 1 AndAlso val2 < 100)
        
        Dim r2 As New Random(42)
        Dim val3 = r2.Next(1, 100)
        Console.WriteLine(val1 = val3) ' Deterministic with same seed
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn system_random_nextdouble() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As New Random()
        Dim val = r.NextDouble()
        Console.WriteLine(val >= 0.0 AndAlso val < 1.0)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
