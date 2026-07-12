use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $expr:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let src = format!(
                r#"Module M
    Sub Main()
        Console.WriteLine({})
    End Sub
End Module
"#,
                $expr
            );
            let out = run_vb(&src);
            assert_eq!(out, vec![super::helpers::dotnet_expected_one($expected)]);
        }
    };
}

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_expr_spec!(
    info_spec_typename_reports_integer_literal,
    r#"TypeName(1)"#,
    "Integer"
);
vb_expr_spec!(
    info_spec_typename_reports_string_literal,
    r#"TypeName("vb")"#,
    "String"
);
vb_expr_spec!(
    info_spec_typename_reports_date_literal,
    r#"TypeName(#5/14/2024#)"#,
    "Date"
);
vb_expr_spec!(
    info_spec_typename_reports_boolean_literal,
    r#"TypeName(True)"#,
    "Boolean"
);
vb_full_spec!(
    info_spec_typename_reports_array_instance,
    r#"Module M
    Sub Main()
        Dim items() As Integer = {1, 2}
        Console.WriteLine(TypeName(items))
    End Sub
End Module"#,
    ["Integer()"]
);
vb_full_spec!(
    info_spec_typename_reports_object_instance,
    r#"Class Box : End Class
Module M
    Sub Main()
        Console.WriteLine(TypeName(New Box()))
    End Sub
End Module"#,
    ["Box"]
);
vb_expr_spec!(
    info_spec_vartype_reports_integer_literal,
    r#"VarType(1)"#,
    "2"
);
vb_expr_spec!(
    info_spec_vartype_reports_string_literal,
    r#"VarType("vb")"#,
    "8"
);
vb_expr_spec!(
    info_spec_vartype_reports_boolean_literal,
    r#"VarType(True)"#,
    "11"
);
vb_expr_spec!(
    info_spec_vartype_reports_date_literal,
    r#"VarType(#5/14/2024#)"#,
    "7"
);
vb_full_spec!(
    info_spec_vartype_reports_array_instance,
    r#"Module M
    Sub Main()
        Dim items() As Integer = {1, 2}
        Console.WriteLine(VarType(items))
    End Sub
End Module"#,
    ["8194"]
);
vb_full_spec!(
    info_spec_isarray_is_true_for_array_variable,
    r#"Module M
    Sub Main()
        Dim items() As Integer = {1, 2}
        Console.WriteLine(IsArray(items))
    End Sub
End Module"#,
    ["true"]
);
vb_full_spec!(
    info_spec_isarray_is_false_for_scalar_variable,
    r#"Module M
    Sub Main()
        Dim value As Integer = 5
        Console.WriteLine(IsArray(value))
    End Sub
End Module"#,
    ["false"]
);
vb_full_spec!(
    info_spec_isobject_is_true_for_class_instance,
    r#"Class Box : End Class
Module M
    Sub Main()
        Console.WriteLine(IsObject(New Box()))
    End Sub
End Module"#,
    ["true"]
);
vb_expr_spec!(
    info_spec_isobject_is_false_for_scalar_value,
    r#"IsObject(5)"#,
    "false"
);
vb_full_spec!(
    info_spec_isreference_is_true_for_class_instance,
    r#"Class Box : End Class
Module M
    Sub Main()
        Console.WriteLine(IsReference(New Box()))
    End Sub
End Module"#,
    ["true"]
);
vb_expr_spec!(
    info_spec_isreference_is_false_for_scalar_value,
    r#"IsReference(5)"#,
    "false"
);
vb_full_spec!(
    info_spec_isnothing_is_true_for_nothing_reference,
    r#"Class Box : End Class
Module M
    Sub Main()
        Dim value As Box = Nothing
        Console.WriteLine(IsNothing(value))
    End Sub
End Module"#,
    ["true"]
);
vb_full_spec!(
    info_spec_isnothing_is_false_for_non_nothing_reference,
    r#"Class Box : End Class
Module M
    Sub Main()
        Dim value As New Box()
        Console.WriteLine(IsNothing(value))
    End Sub
End Module"#,
    ["false"]
);
vb_expr_spec!(
    info_spec_choose_can_return_first_branch,
    r#"Choose(1, "a", "b", "c")"#,
    "a"
);
vb_expr_spec!(
    info_spec_choose_can_return_third_branch,
    r#"Choose(3, "a", "b", "c")"#,
    "c"
);
vb_expr_spec!(
    info_spec_choose_can_return_string_value,
    r#"Choose(2, "left", "right")"#,
    "right"
);
vb_expr_spec!(
    info_spec_switch_returns_first_true_branch,
    r#"Switch(False, "x", True, "y")"#,
    "y"
);
vb_expr_spec!(
    info_spec_switch_returns_later_true_branch,
    r#"Switch(False, "x", False, "y", True, "z")"#,
    "z"
);
vb_expr_spec!(
    info_spec_switch_returns_nothing_when_no_branch_matches,
    r#"IsNothing(Switch(False, "x"))"#,
    "true"
);
vb_expr_spec!(
    info_spec_iif_returns_true_branch_text,
    r#"IIf(True, "yes", "no")"#,
    "yes"
);
vb_expr_spec!(
    info_spec_iif_returns_false_branch_text,
    r#"IIf(False, "yes", "no")"#,
    "no"
);
vb_expr_spec!(
    info_spec_iif_can_return_numeric_branch,
    r#"IIf(True, 7, 9)"#,
    "7"
);
vb_full_spec!(
    info_spec_iif_evaluates_both_branches_before_selection,
    r#"Module M
    Function LeftValue() As String
        Console.WriteLine("left")
        Return "yes"
    End Function
    Function RightValue() As String
        Console.WriteLine("right")
        Return "no"
    End Function
    Sub Main()
        Console.WriteLine(IIf(True, LeftValue(), RightValue()))
    End Sub
End Module"#,
    ["left", "right", "yes"]
);
vb_expr_spec!(
    info_spec_partition_formats_lower_bucket,
    r#"Partition(3, 0, 9, 5)"#,
    " 0: 4"
);
vb_expr_spec!(
    info_spec_partition_formats_middle_bucket,
    r#"Partition(5, 0, 9, 5)"#,
    " 5: 9"
);
vb_expr_spec!(
    info_spec_partition_formats_upper_bucket,
    r#"Partition(12, 0, 9, 5)"#,
    "10:10"
);
vb_expr_spec!(
    info_spec_timer_returns_nonnegative_value,
    r#"Timer >= 0"#,
    "true"
);
vb_expr_spec!(
    info_spec_command_returns_string_value,
    r#"TypeName(Command())"#,
    "String"
);
vb_expr_spec!(
    info_spec_environ_returns_string_value,
    r#"TypeName(Environ("PATH"))"#,
    "String"
);
vb_expr_spec!(
    info_spec_typename_reports_generic_list_instance,
    r#"TypeName(New List(Of Integer)())"#,
    "List`1"
);
vb_expr_spec!(
    info_spec_typename_reports_dictionary_instance,
    r#"TypeName(New Dictionary(Of String, Integer)())"#,
    "Dictionary`2"
);
vb_full_spec!(
    info_spec_vartype_reports_object_instance,
    r#"Class Box : End Class
Module M
    Sub Main()
        Console.WriteLine(VarType(New Box()))
    End Sub
End Module"#,
    ["9"]
);
vb_full_spec!(
    info_spec_typename_reports_structure_value,
    r#"Structure Point
    Public X As Integer
End Structure
Module M
    Sub Main()
        Dim p As Point
        Console.WriteLine(TypeName(p))
    End Sub
End Module"#,
    ["Point"]
);
vb_full_spec!(
    info_spec_typename_reports_enum_value,
    r#"Enum Tone
    Low
    High
End Enum
Module M
    Sub Main()
        Console.WriteLine(TypeName(Tone.High))
    End Sub
End Module"#,
    ["Tone"]
);
vb_full_spec!(
    info_spec_typename_reports_delegate_value,
    r#"Module M
    Sub Main()
        Dim value As Func(Of Integer, Integer) = Function(x) x + 1
        Console.WriteLine(TypeName(value))
    End Sub
End Module"#,
    ["Func`2"]
);
vb_full_spec!(
    info_spec_isreference_is_true_for_array_instance,
    r#"Module M
    Sub Main()
        Dim items() As Integer = {1, 2}
        Console.WriteLine(IsReference(items))
    End Sub
End Module"#,
    ["true"]
);
vb_full_spec!(
    info_spec_isobject_is_true_for_array_instance,
    r#"Module M
    Sub Main()
        Dim items() As Integer = {1, 2}
        Console.WriteLine(IsObject(items))
    End Sub
End Module"#,
    ["true"]
);
vb_expr_spec!(
    info_spec_isarray_is_true_for_array_function_result,
    r#"IsArray(Array(1, 2, 3))"#,
    "true"
);
vb_expr_spec!(
    info_spec_choose_with_index_two_returns_second_item,
    r#"Choose(2, 10, 20, 30)"#,
    "20"
);
vb_expr_spec!(
    info_spec_switch_can_return_numeric_result,
    r#"Switch(False, 1, True, 2)"#,
    "2"
);
vb_expr_spec!(
    info_spec_typename_reports_nothing_as_nothing,
    r#"TypeName(Nothing)"#,
    "Nothing"
);
vb_expr_spec!(
    info_spec_vartype_reports_empty_variant_for_nothing,
    r#"VarType(Nothing)"#,
    "0"
);
vb_full_spec!(
    info_spec_isreference_can_distinguish_string_variable,
    r#"Module M
    Sub Main()
        Dim value As String = "vb"
        Console.WriteLine(IsReference(value))
    End Sub
End Module"#,
    ["true"]
);
vb_full_spec!(
    info_spec_isobject_can_distinguish_string_variable,
    r#"Module M
    Sub Main()
        Dim value As String = "vb"
        Console.WriteLine(IsObject(value))
    End Sub
End Module"#,
    ["false"]
);
