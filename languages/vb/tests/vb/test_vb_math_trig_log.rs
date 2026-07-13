use super::helpers::run_vb;

#[test]
fn math_trig_log() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Abs
        Console.WriteLine(Abs(-15.5))
        
        ' Sqrt
        Console.WriteLine(Sqrt(16))
        
        ' Trig
        Console.WriteLine(Int(Cos(0)))
        Console.WriteLine(Int(Sin(0)))
        Console.WriteLine(Int(Tan(0)))
        
        ' Log / Exp
        Console.WriteLine(Exp(0))
        Console.WriteLine(Int(Log(Exp(1))))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15.5", "4", "1", "0", "0", "1", "1"]);
}
