use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: MyBase & MyClass Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn mybase_myclass_advanced() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable Sub Print()
        Console.WriteLine("Base Print")
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Overrides Sub Print()
        Console.WriteLine("Derived Print")
    End Sub
    
    Public Sub Test()
        ' MyBase calls the base class implementation regardless of overriding
        MyBase.Print()
        
        ' MyClass calls the implementation in THIS class, skipping further overrides
        MyClass.Print()
    End Sub
End Class

Class MoreDerived
    Inherits Derived
    
    Public Overrides Sub Print()
        Console.WriteLine("MoreDerived Print")
    End Sub
End Class

Module M
    Sub Main()
        Dim md As New MoreDerived()
        md.Test() ' Will print Base Print, then Derived Print (due to MyClass in Derived)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Base Print", "Derived Print"]);
}
