use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Attributes (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn custom_attributes_named_arguments() {
    let out = run_vb(
        r#"
Imports System

<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method)>
Public Class InfoAttribute
    Inherits Attribute
    
    Public Property Author As String
    Public Property Version As String
    
    Public Sub New()
    End Sub
End Class

<Info(Author:="Alice", Version:="1.0")>
Class MyComponent
End Class

Module M
    Sub Main()
        Dim t As Type = GetType(MyComponent)
        Dim attrs() As Object = t.GetCustomAttributes(GetType(InfoAttribute), False)
        Dim info As InfoAttribute = DirectCast(attrs(0), InfoAttribute)
        
        Console.WriteLine(info.Author)
        Console.WriteLine(info.Version)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "1.0"]);
}
