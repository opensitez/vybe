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

vb_case!(
    byref_integer_parameter_updates_caller_value,
    r#"
Module M
    Sub Bump(ByRef value As Integer)
        value = value + 1
    End Sub

    Sub Main()
        Dim x As Integer = 4
        Bump(x)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    ["5"]
);

vb_case!(
    byref_integer_parameter_accumulates_across_multiple_calls,
    r#"
Module M
    Sub Bump(ByRef value As Integer)
        value = value + 3
    End Sub

    Sub Main()
        Dim x As Integer = 2
        Bump(x)
        Bump(x)
        Bump(x)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    ["11"]
);

vb_case!(
    byref_integer_parameter_can_decrement_value,
    r#"
Module M
    Sub Lower(ByRef value As Integer)
        value = value - 2
    End Sub

    Sub Main()
        Dim x As Integer = 9
        Lower(x)
        Lower(x)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    ["5"]
);

vb_case!(
    byref_integer_parameter_can_assign_absolute_value,
    r#"
Module M
    Sub ResetValue(ByRef value As Integer)
        value = 42
    End Sub

    Sub Main()
        Dim x As Integer = 3
        ResetValue(x)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    ["42"]
);

vb_case!(
    byref_string_parameter_updates_original_variable,
    r#"
Module M
    Sub AppendSuffix(ByRef text As String)
        text = text & "-done"
    End Sub

    Sub Main()
        Dim text As String = "task"
        AppendSuffix(text)
        Console.WriteLine(text)
    End Sub
End Module
"#,
    ["task-done"]
);

vb_case!(
    byref_string_parameter_can_replace_entire_value,
    r#"
Module M
    Sub ReplaceText(ByRef text As String)
        text = "replaced"
    End Sub

    Sub Main()
        Dim text As String = "old"
        ReplaceText(text)
        Console.WriteLine(text)
    End Sub
End Module
"#,
    ["replaced"]
);

vb_case!(
    byref_boolean_parameter_can_flip_flag,
    r#"
Module M
    Sub Toggle(ByRef flag As Boolean)
        flag = Not flag
    End Sub

    Sub Main()
        Dim flag As Boolean = False
        Toggle(flag)
        Console.WriteLine(flag)
    End Sub
End Module
    "#,
    ["true"]
);

vb_case!(
    byref_boolean_parameter_can_flip_multiple_times,
    r#"
Module M
    Sub Toggle(ByRef flag As Boolean)
        flag = Not flag
    End Sub

    Sub Main()
        Dim flag As Boolean = True
        Toggle(flag)
        Toggle(flag)
        Console.WriteLine(flag)
    End Sub
End Module
    "#,
    ["true"]
);

vb_case!(
    byref_parameter_mutation_inside_loop_persists_to_caller,
    r#"
Module M
    Sub AddRange(ByRef value As Integer)
        For i As Integer = 1 To 4
            value = value + i
        Next
    End Sub

    Sub Main()
        Dim total As Integer = 0
        AddRange(total)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    byref_parameter_can_be_forwarded_to_nested_helper,
    r#"
Module M
    Sub Inner(ByRef value As Integer)
        value = value * 2
    End Sub

    Sub Outer(ByRef value As Integer)
        Inner(value)
    End Sub

    Sub Main()
        Dim value As Integer = 6
        Outer(value)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    ["12"]
);

vb_case!(
    byref_parameter_updates_caller_when_used_with_if_branch,
    r#"
Module M
    Sub Adjust(ByRef value As Integer, shouldBoost As Boolean)
        If shouldBoost Then
            value = value + 10
        Else
            value = value + 1
        End If
    End Sub

    Sub Main()
        Dim x As Integer = 5
        Adjust(x, True)
        Adjust(x, False)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    ["16"]
);

vb_case!(
    byref_parameter_can_swap_two_variables,
    r#"
Module M
    Sub Swap(ByRef left As Integer, ByRef right As Integer)
        Dim temp As Integer = left
        left = right
        right = temp
    End Sub

    Sub Main()
        Dim a As Integer = 3
        Dim b As Integer = 8
        Swap(a, b)
        Console.WriteLine(a)
        Console.WriteLine(b)
    End Sub
End Module
"#,
    ["8", "3"]
);

vb_case!(
    byref_parameter_can_mutate_array_element,
    r#"
Module M
    Sub Bump(ByRef value As Integer)
        value = value + 5
    End Sub

    Sub Main()
        Dim values() As Integer = {1, 2, 3}
        Bump(values(1))
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    ["7"]
);

vb_case!(
    byref_parameter_can_update_module_level_variable,
    r#"
Module M
    Dim total As Integer = 10

    Sub Bump(ByRef value As Integer)
        value = value + 4
    End Sub

    Sub Main()
        Bump(total)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["14"]
);

vb_case!(
    byref_parameter_can_remove_one_matching_substring,
    r#"
Module M
    Sub TrimBang(ByRef text As String)
        text = text.Replace("!", "")
    End Sub

    Sub Main()
        Dim text As String = "Hi!!!"
        TrimBang(text)
        Console.WriteLine(text)
    End Sub
End Module
    "#,
    ["Hi!!"]
);

vb_case!(
    byref_parameter_can_assign_nothing_to_string,
    r#"
Module M
    Sub ClearText(ByRef text As String)
        text = Nothing
    End Sub

    Sub Main()
        Dim text As String = "value"
        ClearText(text)
        Console.WriteLine(IsNothing(text))
    End Sub
End Module
    "#,
    ["true"]
);

vb_case!(
    byref_parameter_can_build_running_total_with_varying_deltas,
    r#"
Module M
    Sub Add(ByRef total As Integer, amount As Integer)
        total = total + amount
    End Sub

    Sub Main()
        Dim total As Integer = 0
        Add(total, 3)
        Add(total, 7)
        Add(total, -2)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["8"]
);

vb_case!(
    byref_parameter_can_update_value_from_function_result,
    r#"
Module M
    Function NextValue() As Integer
        Return 21
    End Function

    Sub Replace(ByRef value As Integer)
        value = NextValue()
    End Sub

    Sub Main()
        Dim current As Integer = 1
        Replace(current)
        Console.WriteLine(current)
    End Sub
End Module
"#,
    ["21"]
);

vb_case!(
    byref_parameter_can_compose_string_prefix_and_suffix,
    r#"
Module M
    Sub Decorate(ByRef text As String)
        text = "[" & text & "]"
    End Sub

    Sub Main()
        Dim text As String = "core"
        Decorate(text)
        Console.WriteLine(text)
    End Sub
End Module
"#,
    ["[core]"]
);

vb_case!(
    byref_parameter_can_mutate_integer_twice_in_same_call,
    r#"
Module M
    Sub Adjust(ByRef value As Integer)
        value = value + 2
        value = value * 3
    End Sub

    Sub Main()
        Dim value As Integer = 4
        Adjust(value)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    ["18"]
);
