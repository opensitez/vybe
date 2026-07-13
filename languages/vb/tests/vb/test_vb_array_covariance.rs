use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Covariance
// ═══════════════════════════════════════════════════════════

#[test]
fn array_covariance() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable Sub Show()
        Console.WriteLine("Base")
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Overrides Sub Show()
        Console.WriteLine("Derived")
    End Sub
End Class

Module M
    Sub Main()
        ' In VB.NET arrays of reference types are covariant
        Dim derivedArr() As Derived = { New Derived(), New Derived() }
        Dim baseArr() As Base = derivedArr
        
        For Each b In baseArr
            b.Show()
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Derived", "Derived"]);
}
