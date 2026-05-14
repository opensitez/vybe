use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_case!(directcast_reads_boxed_integer_value, r#"
Module M
    Sub Main()
        Dim boxed As Object = 12
        Console.WriteLine(DirectCast(boxed, Integer))
    End Sub
End Module
"#, ["12"]);

vb_case!(directcast_reads_boxed_string_value, r#"
Module M
    Sub Main()
        Dim boxed As Object = "hello"
        Console.WriteLine(DirectCast(boxed, String))
    End Sub
End Module
"#, ["hello"]);

vb_case!(directcast_reads_boxed_boolean_value, r#"
Module M
    Sub Main()
        Dim boxed As Object = True
        Console.WriteLine(DirectCast(boxed, Boolean))
    End Sub
End Module
"#, ["True"]);

vb_case!(trycast_returns_same_string_reference_when_types_match, r#"
Module M
    Sub Main()
        Dim boxed As Object = "sample"
        Dim value As String = TryCast(boxed, String)
        Console.WriteLine(value)
    End Sub
End Module
"#, ["sample"]);

vb_case!(trycast_returns_nothing_for_incompatible_reference_type, r#"
Module M
    Class Greeter
    End Class

    Sub Main()
        Dim boxed As Object = New Greeter()
        Dim value As String = TryCast(boxed, String)
        Console.WriteLine(IsNothing(value))
    End Sub
End Module
"#, ["True"]);

vb_case!(directcast_supports_custom_reference_type, r#"
Module M
    Class Greeter
        Public Message As String = "hi"
    End Class

    Sub Main()
        Dim boxed As Object = New Greeter()
        Dim value As Greeter = DirectCast(boxed, Greeter)
        Console.WriteLine(value.Message)
    End Sub
End Module
"#, ["hi"]);

vb_case!(trycast_supports_custom_reference_type, r#"
Module M
    Class Greeter
        Public Message As String = "hello"
    End Class

    Sub Main()
        Dim boxed As Object = New Greeter()
        Dim value As Greeter = TryCast(boxed, Greeter)
        Console.WriteLine(value.Message)
    End Sub
End Module
"#, ["hello"]);

vb_case!(ctype_converts_integer_to_string, r#"
Module M
    Sub Main()
        Console.WriteLine(CType(42, String))
    End Sub
End Module
"#, ["42"]);

vb_case!(ctype_converts_string_to_integer, r#"
Module M
    Sub Main()
        Console.WriteLine(CType("55", Integer))
    End Sub
End Module
"#, ["55"]);

vb_case!(cint_rounds_numeric_value_to_integer, r#"
Module M
    Sub Main()
        Console.WriteLine(CInt(7.0))
    End Sub
End Module
"#, ["7"]);

vb_case!(cdbl_converts_integer_to_double_shape, r#"
Module M
    Sub Main()
        Console.WriteLine(CDbl(5))
    End Sub
End Module
"#, ["5"]);

vb_case!(cstr_converts_boolean_to_text, r#"
Module M
    Sub Main()
        Console.WriteLine(CStr(False))
    End Sub
End Module
"#, ["False"]);

vb_case!(cbool_converts_boolean_literal, r#"
Module M
    Sub Main()
        Console.WriteLine(CBool(True))
    End Sub
End Module
"#, ["True"]);

vb_case!(directcast_can_be_used_inside_expression, r#"
Module M
    Sub Main()
        Dim boxed As Object = 9
        Console.WriteLine(DirectCast(boxed, Integer) + 1)
    End Sub
End Module
"#, ["10"]);

vb_case!(trycast_result_can_flow_through_isnothing_check, r#"
Module M
    Sub Main()
        Dim boxed As Object = "vb"
        Dim value As String = TryCast(boxed, String)
        If IsNothing(value) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine(value)
        End If
    End Sub
End Module
"#, ["vb"]);

vb_case!(directcast_handles_negative_integer_values, r#"
Module M
    Sub Main()
        Dim boxed As Object = -8
        Console.WriteLine(DirectCast(boxed, Integer))
    End Sub
End Module
"#, ["-8"]);

vb_case!(directcast_handles_large_integer_values, r#"
Module M
    Sub Main()
        Dim boxed As Object = 2048
        Console.WriteLine(DirectCast(boxed, Integer))
    End Sub
End Module
"#, ["2048"]);

vb_case!(cstr_converts_integer_expression_result, r#"
Module M
    Sub Main()
        Console.WriteLine(CStr(20 + 22))
    End Sub
End Module
"#, ["42"]);

vb_case!(ctype_can_convert_boxed_integer_back_to_integer, r#"
Module M
    Sub Main()
        Dim boxed As Object = 18
        Console.WriteLine(CType(boxed, Integer))
    End Sub
End Module
"#, ["18"]);

vb_case!(trycast_on_nothing_reference_returns_nothing, r#"
Module M
    Sub Main()
        Dim boxed As Object = Nothing
        Dim value As String = TryCast(boxed, String)
        Console.WriteLine(IsNothing(value))
    End Sub
End Module
"#, ["True"]);