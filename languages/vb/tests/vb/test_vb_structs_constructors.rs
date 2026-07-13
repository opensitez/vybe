use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Struct Constructors
// ═══════════════════════════════════════════════════════════

#[test]
fn struct_constructor_initialization() {
    let out = run_vb(
        r#"
Structure Size
    Public Width As Integer
    Public Height As Integer
    
    ' Parameterized constructor
    Public Sub New(w As Integer, h As Integer)
        Width = w
        Height = h
    End Sub
End Structure

Module M
    Sub Main()
        Dim s As New Size(1024, 768)
        Console.WriteLine(s.Width)
        Console.WriteLine(s.Height)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1024", "768"]);
}

#[test]
fn struct_default_values_without_new() {
    let out = run_vb(
        r#"
Structure Status
    Public Code As Integer
    Public Message As String
End Structure

Module M
    Sub Main()
        ' A struct can be used without New, fields are zeroed/Nothing
        Dim s As Status
        Console.WriteLine(s.Code)
        Console.WriteLine(IsNothing(s.Message))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "True"]);
}
