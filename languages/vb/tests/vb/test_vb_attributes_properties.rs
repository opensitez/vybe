use super::helpers::run_vb;

#[test]
fn attributes_properties() {
    let out = run_vb(
        r#"
Class Data
    ' Attributes on properties
    <System.Obsolete("Use NewId instead")>
    Public Property Id As Integer
    
    Public Property NewId As Integer
End Class

Module M
    Sub Main()
        Dim d As New Data()
        d.Id = 10
        Console.WriteLine(d.Id)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10"]);
}
