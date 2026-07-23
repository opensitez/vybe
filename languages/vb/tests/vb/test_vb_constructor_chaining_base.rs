use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Constructor Chaining (Me.New & MyBase.New)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_constructor_me_new_chaining() {
    let src = r#"
Class Account
    Public Property Name As String
    Public Property Balance As Decimal

    Public Sub New()
        Me.New("Default", 0D)
    End Sub

    Public Sub New(name As String)
        Me.New(name, 100D)
    End Sub

    Public Sub New(name As String, balance As Decimal)
        Me.Name = name
        Me.Balance = balance
    End Sub
End Class

Module Program
    Sub Main()
        Dim a1 As New Account()
        Dim a2 As New Account("Alice")
        Console.WriteLine(a1.Name & ":" & a1.Balance)
        Console.WriteLine(a2.Name & ":" & a2.Balance)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Default:0", "Alice:100"]);
}

#[test]
fn test_vb_constructor_mybase_new_chaining() {
    let src = r#"
Class BaseResource
    Public Shared InitLog As String = ""
    Public Sub New(msg As String)
        InitLog &= "Base:" & msg & ";"
    End Sub
End Class

Class DerivedResource
    Inherits BaseResource
    Public Sub New(msg As String)
        MyBase.New(msg)
        InitLog &= "Derived:" & msg & ";"
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New DerivedResource("Test")
        Console.WriteLine(BaseResource.InitLog)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Base:Test;Derived:Test;"]);
}

#[test]
fn test_vb_constructor_execution_order_fields_and_ctors() {
    let src = r#"
Class BaseObj
    Public Field1 As String = "BaseField"
    Public Sub New()
        Console.WriteLine("BaseCtor")
    End Sub
End Class

Class DerivedObj
    Inherits BaseObj
    Public Field2 As String = "DerivedField"
    Public Sub New()
        MyBase.New()
        Console.WriteLine("DerivedCtor")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New DerivedObj()
        Console.WriteLine(d.Field1 & ":" & d.Field2)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["BaseCtor", "DerivedCtor", "BaseField:DerivedField"]
    );
}

#[test]
fn test_vb_constructor_protected_accessibility() {
    let src = r#"
Class ProtectedBase
    Protected Sub New()
        Console.WriteLine("Protected Ctor Called")
    End Sub
End Class

Class PublicDerived
    Inherits ProtectedBase
    Public Sub New()
        MyBase.New()
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New PublicDerived()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Protected Ctor Called"]);
}
