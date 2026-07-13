use super::helpers::run_vb;

#[test]
fn widening_conversions() {
    let out = run_vb(
        r#"
Option Strict On

Class Base
End Class

Class Derived
    Inherits Base
End Class

Module M
    Sub Main()
        ' Widening numeric conversion
        Dim i As Integer = 10
        Dim d As Double = i
        Console.WriteLine(d)
        
        ' Widening reference conversion
        Dim dev As New Derived()
        Dim b As Base = dev
        Console.WriteLine(b IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "True"]);
}
