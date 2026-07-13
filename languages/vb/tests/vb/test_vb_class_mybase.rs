use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Classes (MyBase and MyClass)
// ═══════════════════════════════════════════════════════════

#[test]
fn class_mybase_method_call() {
    let out = run_vb(
        r#"
Class Person
    Public Overridable Function Greet() As String
        Return "Hello"
    End Function
End Class

Class Employee
    Inherits Person
    
    Public Overrides Function Greet() As String
        Return MyBase.Greet() & " Boss"
    End Function
End Class

Module M
    Sub Main()
        Dim e As New Employee()
        Console.WriteLine(e.Greet())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello Boss"]);
}

#[test]
fn class_mybase_constructor() {
    let out = run_vb(
        r#"
Class BaseObj
    Public ID As Integer
    Public Sub New(id As Integer)
        Me.ID = id
    End Sub
End Class

Class DerivedObj
    Inherits BaseObj
    Public Name As String
    
    Public Sub New(id As Integer, name As String)
        MyBase.New(id)
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New DerivedObj(42, "Test")
        Console.WriteLine(d.ID)
        Console.WriteLine(d.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Test"]);
}

#[test]
fn class_myclass_keyword() {
    let out = run_vb(
        r#"
Class BasePrinter
    Public Overridable Function GetName() As String
        Return "Base"
    End Function
    
    Public Function PrintName() As String
        ' MyClass forces call to this class's implementation, ignoring overrides
        Return MyClass.GetName()
    End Function
    
    Public Function PrintNamePolymorphic() As String
        Return Me.GetName()
    End Function
End Class

Class DerivedPrinter
    Inherits BasePrinter
    
    Public Overrides Function GetName() As String
        Return "Derived"
    End Function
End Class

Module M
    Sub Main()
        Dim d As New DerivedPrinter()
        Console.WriteLine(d.PrintName())
        Console.WriteLine(d.PrintNamePolymorphic())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Base", "Derived"]);
}
