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

vb_case!(
    byref_can_mutate_structure_fields_in_place,
    r#"
Module M
    Structure Point
        Public X As Integer
        Public Y As Integer
    End Structure

    Sub Move(ByRef point As Point)
        point.X += 1
        point.Y += 2
    End Sub

    Sub Main()
        Dim point As Point
        point.X = 2
        point.Y = 3
        Move(point)
        Console.WriteLine(point.X.ToString() & "," & point.Y.ToString())
    End Sub
End Module
"#,
    ["3,5"]
);

vb_case!(
    byref_can_rebind_reference_type_variable_in_caller,
    r#"
Module M
    Class Holder
        Public Value As String
        Public Sub New(value As String)
            Me.Value = value
        End Sub
    End Class

    Sub Replace(ByRef holder As Holder)
        holder = New Holder("replaced")
    End Sub

    Sub Main()
        Dim value As Holder = New Holder("initial")
        Replace(value)
        Console.WriteLine(value.Value)
    End Sub
End Module
"#,
    ["replaced"]
);

vb_case!(
    byref_reference_type_mutation_keeps_same_instance_without_rebinding,
    r#"
Module M
    Class Holder
        Public Value As Integer
        Public Sub New(value As Integer)
            Me.Value = value
        End Sub
    End Class

    Sub Boost(ByRef holder As Holder)
        holder.Value = holder.Value + 3
    End Sub

    Sub Main()
        Dim value As Holder = New Holder(7)
        Boost(value)
        Console.WriteLine(value.Value)
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    byref_can_rebind_array_variable_not_only_element,
    r#"
Module M
    Sub ReplaceNumbers(ByRef values() As Integer)
        values = New Integer() {4, 5, 6}
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {1, 2, 3}
        ReplaceNumbers(values)
        Console.WriteLine(values.Length)
        Console.WriteLine(values(0) + values(2))
    End Sub
End Module
"#,
    ["3", "10"]
);

vb_case!(
    byref_can_increment_single_array_element_through_alias,
    r#"
Module M
    Sub Increment(ByRef item As Integer)
        item += 1
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {1, 2, 3}
        Increment(values(1))
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    ["3"]
);

vb_case!(
    byref_swaps_array_elements_via_two_aliases,
    r#"
Module M
    Sub Swap(ByRef left As Integer, ByRef right As Integer)
        Dim saved As Integer = left
        left = right
        right = saved
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {8, 1}
        Swap(values(0), values(1))
        Console.WriteLine(values(0))
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    ["1", "8"]
);

vb_case!(
    byref_decimal_parameter_can_change_integral_shape,
    r#"
Module M
    Sub SetTotal(ByRef total As Decimal)
        total = total + 20D
    End Sub

    Sub Main()
        Dim total As Decimal = CDec(22)
        SetTotal(total)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["42"]
);

vb_case!(
    byref_datetime_day_can_advance_one_and_preserve_caller,
    r#"
Module M
    Sub Advance(ByRef value As Date)
        value = value.AddDays(1)
    End Sub

    Sub Main()
        Dim value As Date = New Date(2024, 7, 1)
        Advance(value)
        Console.WriteLine(value.Day.ToString())
    End Sub
End Module
"#,
    ["2"]
);

vb_case!(
    byref_by_reference_nullable_integer_can_turn_to_nothing,
    r#"
Module M
    Sub Clear(ByRef value As Integer?)
        value = Nothing
    End Sub

    Sub Main()
        Dim value As Integer? = 99
        Clear(value)
        Console.WriteLine(IsNothing(value))
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    byref_side_effect_survives_exception_path,
    r#"
Module M
    Sub UpdateThenThrow(ByRef value As Integer)
        value = value + 4
        Throw New Exception("failed")
    End Sub

    Sub Main()
        Dim value As Integer = 6
        Try
            UpdateThenThrow(value)
        Catch ex As Exception
            Console.WriteLine(value)
        End Try
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    byref_nested_dispatch_chain_uses_same_alias_in_multiple_methods,
    r#"
Module M
    Sub Multiply(ByRef value As Integer, multiplier As Integer)
        value *= multiplier
    End Sub

    Sub Apply(ByRef value As Integer)
        Multiply(value, 2)
        Multiply(value, 3)
    End Sub

    Sub Main()
        Dim value As Integer = 4
        Apply(value)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    ["24"]
);
