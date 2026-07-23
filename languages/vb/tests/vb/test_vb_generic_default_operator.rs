use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: CType(Nothing, T) / Default Value Operator
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_default_value_primitives_and_references() {
    let src = r#"
Module Helper
    Public Function GetDefault(Of T)() As T
        Return CType(Nothing, T)
    End Function
End Module

Module Program
    Sub Main()
        Dim defaultInt As Integer = Helper.GetDefault(Of Integer)()
        Dim defaultBool As Boolean = Helper.GetDefault(Of Boolean)()
        Dim defaultStr As String = Helper.GetDefault(Of String)()
        Console.WriteLine(defaultInt)
        Console.WriteLine(defaultBool)
        Console.WriteLine(defaultStr Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "False", "True"]);
}

#[test]
fn test_vb_generic_default_value_nullable() {
    let src = r#"
Module Helper
    Public Function GetDefault(Of T)() As T
        Return CType(Nothing, T)
    End Function
End Module

Module Program
    Sub Main()
        Dim defaultNullable As Nullable(Of Integer) = Helper.GetDefault(Of Nullable(Of Integer))()
        Console.WriteLine(defaultNullable.HasValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}
