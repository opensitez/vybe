use super::helpers::run_vb;

#[test]
fn system_reflection_basic() {
    let out = run_vb(
        r#"
Imports System.Reflection

Class Person
    Public Property Name As String
    Public Property Age As Integer
    
    Public Sub SayHello()
        Console.WriteLine("Hello")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Person()
        Dim t As Type = p.GetType()
        
        Console.WriteLine(t.Name)
        
        Dim props = t.GetProperties()
        Console.WriteLine(props.Length)
        
        Dim m = t.GetMethod("SayHello")
        If m IsNot Nothing Then
            m.Invoke(p, Nothing)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Person", "2", "Hello"]);
}
