use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Attributes Declaration & Reflection Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_attribute_read_properties() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method)>
Public Class InfoAttribute
    Inherits Attribute

    Public ReadOnly Description As String
    Public Property Version As Integer = 1

    Public Sub New(desc As String)
        Me.Description = desc
    End Sub
End Class

<Info("Test Class", Version := 2)>
Class Sample
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(Sample)
        Dim attrs = t.GetCustomAttributes(GetType(InfoAttribute), False)
        If attrs.Length > 0 Then
            Dim info As InfoAttribute = CType(attrs(0), InfoAttribute)
            Console.WriteLine(info.Description & ":V" & info.Version)
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Test Class:V2"]);
}
