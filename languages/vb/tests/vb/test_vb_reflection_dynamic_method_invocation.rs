use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection Dynamic Method Invocation (MethodInfo.Invoke)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_invoke_instance_method() {
    let src = r#"
Imports System.Reflection

Class Calculator
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New Calculator()
        Dim t As Type = calc.GetType()
        Dim mi As MethodInfo = t.GetMethod("Add")
        Dim res As Object = mi.Invoke(calc, New Object() {5, 10})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_reflection_invoke_static_method() {
    let src = r#"
Imports System.Reflection

Class Utils
    Public Shared Function Greet(name As String) As String
        Return "Hello " & name
    End Function
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(Utils)
        Dim mi As MethodInfo = t.GetMethod("Greet")
        Dim res As Object = mi.Invoke(Nothing, New Object() {"World"})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}
