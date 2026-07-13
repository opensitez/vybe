use super::helpers::run_vb;

#[test]
fn readonly_struct_fields() {
    let out = run_vb(
        r#"
Structure Point
    Public ReadOnly X As Integer
    Public ReadOnly Y As Integer
    
    Public Sub New(xVal As Integer, yVal As Integer)
        X = xVal
        Y = yVal
    End Sub
End Structure

Module M
    Sub Main()
        Dim p As New Point(10, 20)
        Console.WriteLine(p.X)
        Console.WriteLine(p.Y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}
