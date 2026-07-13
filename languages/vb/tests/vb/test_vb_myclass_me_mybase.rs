use super::helpers::run_vb;

#[test]
fn myclass_me_mybase() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable Sub Print()
        Console.WriteLine("Base")
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Overrides Sub Print()
        Console.WriteLine("Derived")
    End Sub
    
    Public Sub Test()
        ' Me calls the most derived override (Derived)
        Me.Print()
        
        ' MyClass calls the implementation in the current class, ignoring overrides
        ' Wait, MyClass calls the method as if it were not virtual, 
        ' so MyClass.Print() in Derived calls Derived.Print(), but if Derived were inherited and Print overridden again, MyClass.Print() in Derived would still call Derived.Print().
        MyClass.Print()
        
        ' MyBase calls the base class implementation
        MyBase.Print()
    End Sub
End Class

Class MoreDerived
    Inherits Derived
    
    Public Overrides Sub Print()
        Console.WriteLine("MoreDerived")
    End Sub
End Class

Module M
    Sub Main()
        Dim obj As New MoreDerived()
        obj.Test()
    End Sub
End Module
"#,
    );
    // When obj.Test() is called (defined in Derived):
    // Me.Print() calls MoreDerived.Print() -> "MoreDerived"
    // MyClass.Print() calls Derived.Print() -> "Derived"
    // MyBase.Print() inside Derived calls Base.Print() -> "Base"
    assert_eq!(out, vec!["MoreDerived", "Derived", "Base"]);
}
