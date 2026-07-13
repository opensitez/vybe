use super::helpers::run_vb;

#[test]
fn using_multiple_vars() {
    let out = run_vb(
        r#"
Imports System

Class Resource
    Implements IDisposable
    
    Public Name As String
    
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine(Name & " Disposed")
    End Sub
End Class

Module M
    Sub Main()
        ' Multiple variables of the same type in one Using statement
        Using r1 As New Resource() With {.Name = "R1"}, r2 As New Resource() With {.Name = "R2"}
            Console.WriteLine("Inside Using")
        End Using
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inside Using", "R2 Disposed", "R1 Disposed"]);
}
