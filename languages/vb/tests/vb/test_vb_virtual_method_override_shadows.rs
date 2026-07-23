use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Overridable, Overrides & Shadows Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_overridable_polymorphic_dispatch() {
    let src = r#"
Class Animal
    Public Overridable Sub Speak()
        Console.WriteLine("Animal sound")
    End Sub
End Class

Class Dog
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("Woof")
    End Sub
End Class

Module Program
    Sub Main()
        Dim a As Animal = New Dog()
        a.Speak()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Woof"]);
}

#[test]
fn test_vb_shadows_static_binding_by_type() {
    let src = r#"
Class Parent
    Public Sub Show()
        Console.WriteLine("Parent Show")
    End Sub
End Class

Class Child
    Inherits Parent
    Public Shadows Sub Show()
        Console.WriteLine("Child Show")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Child()
        Dim p As Parent = c
        c.Show()
        p.Show()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child Show", "Parent Show"]);
}

#[test]
fn test_vb_notoverridable_prevents_further_override() {
    let src = r#"
Class BaseClass
    Public Overridable Sub Action()
        Console.WriteLine("Base")
    End Sub
End Class

Class MidClass
    Inherits BaseClass
    Public NotOverridable Overrides Sub Action()
        Console.WriteLine("Mid")
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As BaseClass = New MidClass()
        b.Action()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Mid"]);
}

#[test]
fn test_vb_shadows_overloaded_member_by_name() {
    let src = r#"
Class BasePrinter
    Public Sub Print(x As Integer)
        Console.WriteLine("Base Int: " & x)
    End Sub
End Class

Class DerivedPrinter
    Inherits BasePrinter
    Public Shadows Sub Print(s As String)
        Console.WriteLine("Derived String: " & s)
    End Sub
End Class

Module Program
    Sub Main()
        Dim dp As New DerivedPrinter()
        dp.Print("Hello")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Derived String: Hello"]);
}

#[test]
fn test_vb_mybase_call_virtual_method() {
    let src = r#"
Class BaseService
    Public Overridable Sub Execute()
        Console.WriteLine("Base Service")
    End Sub
End Class

Class ExtendedService
    Inherits BaseService
    Public Overrides Sub Execute()
        MyBase.Execute()
        Console.WriteLine("Extended Service")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As BaseService = New ExtendedService()
        s.Execute()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Base Service", "Extended Service"]);
}
