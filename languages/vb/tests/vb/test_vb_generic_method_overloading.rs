use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Method Overloading & Type Inference
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_method_overload_non_generic_preference() {
    let src = r#"
Module Utility
    Public Sub Display(Of T)(val As T)
        Console.WriteLine("Generic: " & val.ToString())
    End Sub

    Public Sub Display(val As String)
        Console.WriteLine("NonGenericString: " & val)
    End Sub
End Module

Module Program
    Sub Main()
        Utility.Display("Hello")
        Utility.Display(123)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NonGenericString: Hello", "Generic: 123"]);
}

#[test]
fn test_vb_generic_method_overload_by_arity() {
    let src = r#"
Module Converter
    Public Function ConvertVal(Of T)(val As Object) As T
        Return CType(val, T)
    End Function

    Public Function ConvertVal(Of T1, T2)(val1 As Object, val2 As Object) As String
        Return val1.ToString() & "-" & val2.ToString()
    End Function
End Module

Module Program
    Sub Main()
        Dim i As Integer = Converter.ConvertVal(Of Integer)("100")
        Dim s As String = Converter.ConvertVal(Of Integer, String)(1, 2)
        Console.WriteLine(i)
        Console.WriteLine(s)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100", "1-2"]);
}

#[test]
fn test_vb_generic_method_type_inference_from_args() {
    let src = r#"
Module Helper
    Public Function Identity(Of T)(item As T) As T
        Return item
    End Function
End Module

Module Program
    Sub Main()
        Dim resStr = Helper.Identity("InferString")
        Dim resInt = Helper.Identity(42)
        Console.WriteLine(resStr)
        Console.WriteLine(resInt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InferString", "42"]);
}
