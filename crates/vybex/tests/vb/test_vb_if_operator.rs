use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_case!(if_operator_returns_true_branch_for_boolean_condition, r#"
Module M
    Sub Main()
        Console.WriteLine(If(True, 1, 9))
    End Sub
End Module
"#, ["1"]);

vb_case!(if_operator_returns_false_branch_for_boolean_condition, r#"
Module M
    Sub Main()
        Console.WriteLine(If(False, 1, 9))
    End Sub
End Module
"#, ["9"]);

vb_case!(if_operator_can_use_comparison_expression_as_condition, r#"
Module M
    Sub Main()
        Console.WriteLine(If(3 < 5, 10, 20))
    End Sub
End Module
"#, ["10"]);

vb_case!(if_operator_can_choose_between_string_values, r#"
Module M
    Sub Main()
        Console.WriteLine(If(2 > 4, "left", "right"))
    End Sub
End Module
"#, ["right"]);

vb_case!(if_operator_can_select_function_call_result, r#"
Module M
    Function AddOne(value As Integer) As Integer
        Return value + 1
    End Function

    Sub Main()
        Console.WriteLine(If(True, AddOne(4), AddOne(9)))
    End Sub
End Module
"#, ["5"]);

vb_case!(if_operator_can_nest_inside_false_branch, r#"
Module M
    Sub Main()
        Console.WriteLine(If(False, 1, If(True, 2, 3)))
    End Sub
End Module
"#, ["2"]);

vb_case!(if_operator_coalesces_nothing_to_fallback_string, r#"
Module M
    Sub Main()
        Dim value As String = Nothing
        Console.WriteLine(If(value, "fallback"))
    End Sub
End Module
"#, ["fallback"]);

vb_case!(if_operator_preserves_existing_string_value, r#"
Module M
    Sub Main()
        Dim value As String = "alpha"
        Console.WriteLine(If(value, "fallback"))
    End Sub
End Module
"#, ["alpha"]);

vb_case!(if_operator_treats_empty_string_as_non_nothing_value, r#"
Module M
    Sub Main()
        Dim value As String = ""
        Console.WriteLine("[" & If(value, "fallback") & "]")
    End Sub
End Module
"#, ["[]"]);

vb_case!(if_operator_can_coalesce_function_result_that_returns_nothing, r#"
Module M
    Function MaybeText(flag As Boolean) As String
        If flag Then
            Return "present"
        End If
        Return Nothing
    End Function

    Sub Main()
        Console.WriteLine(If(MaybeText(False), "missing"))
    End Sub
End Module
"#, ["missing"]);