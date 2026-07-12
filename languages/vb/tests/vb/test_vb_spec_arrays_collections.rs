use super::helpers::run_vb;

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_full_spec!(
    array_spec_array_function_returns_variant_array,
    r#"Module M : Sub Main() : Dim items = Array(1, 2, 3) : Console.WriteLine(items(0)) : Console.WriteLine(items(2)) : End Sub : End Module"#,
    ["1", "3"]
);
vb_full_spec!(
    array_spec_ubound_reports_last_index_for_array_function_result,
    r#"Module M : Sub Main() : Dim items = Array("a", "b", "c") : Console.WriteLine(UBound(items)) : End Sub : End Module"#,
    ["2"]
);
vb_full_spec!(
    array_spec_lbound_reports_zero_for_array_function_result,
    r#"Module M : Sub Main() : Dim items = Array("a", "b", "c") : Console.WriteLine(LBound(items)) : End Sub : End Module"#,
    ["0"]
);
vb_full_spec!(
    array_spec_erase_resets_fixed_integer_array_values,
    r#"Module M : Sub Main() : Dim values(2) As Integer : values(0)=1 : values(1)=2 : values(2)=3 : Erase values : Console.WriteLine(values(0)) : Console.WriteLine(values(2)) : End Sub : End Module"#,
    ["0", "0"]
);
vb_full_spec!(
    array_spec_erase_resets_dynamic_string_array_reference,
    r#"Module M : Sub Main() : Dim values() As String = {"a", "b"} : Erase values : Console.WriteLine(IsNothing(values)) : End Sub : End Module"#,
    ["true"]
);
vb_full_spec!(
    array_spec_redim_grows_dynamic_array_without_preserve,
    r#"Module M : Sub Main() : Dim values() As Integer : ReDim values(3) : values(3)=9 : Console.WriteLine(UBound(values)) : Console.WriteLine(values(3)) : End Sub : End Module"#,
    ["3", "9"]
);
vb_full_spec!(
    array_spec_redim_preserve_keeps_existing_prefix_values,
    r#"Module M : Sub Main() : Dim values() As Integer = {1,2} : ReDim Preserve values(3) : Console.WriteLine(values(0)) : Console.WriteLine(values(1)) : End Sub : End Module"#,
    ["1", "2"]
);
vb_full_spec!(
    array_spec_redim_preserve_allows_new_tail_assignment,
    r#"Module M : Sub Main() : Dim values() As Integer = {1,2} : ReDim Preserve values(3) : values(3)=8 : Console.WriteLine(values(3)) : End Sub : End Module"#,
    ["8"]
);
vb_full_spec!(
    array_spec_redim_can_shrink_dynamic_array,
    r#"Module M : Sub Main() : Dim values() As Integer = {1,2,3,4} : ReDim Preserve values(1) : Console.WriteLine(UBound(values)) : End Sub : End Module"#,
    ["1"]
);
vb_full_spec!(
    array_spec_dynamic_array_can_start_unallocated_then_redim,
    r#"Module M : Sub Main() : Dim values() As Integer : ReDim values(0) : values(0)=5 : Console.WriteLine(values(0)) : End Sub : End Module"#,
    ["5"]
);
vb_full_spec!(
    array_spec_dynamic_array_can_redim_to_single_element,
    r#"Module M : Sub Main() : Dim values() As String : ReDim values(0) : values(0)="only" : Console.WriteLine(values(0)) : End Sub : End Module"#,
    ["only"]
);
vb_full_spec!(
    array_spec_multidimensional_array_reads_row_and_column_item,
    r#"Module M : Sub Main() : Dim grid(1,1) As Integer : grid(1,1)=9 : Console.WriteLine(grid(1,1)) : End Sub : End Module"#,
    ["9"]
);
vb_full_spec!(
    array_spec_multidimensional_array_can_sum_all_cells,
    r#"Module M : Sub Main() : Dim grid(1,1) As Integer : grid(0,0)=1 : grid(0,1)=2 : grid(1,0)=3 : grid(1,1)=4 : Console.WriteLine(grid(0,0)+grid(0,1)+grid(1,0)+grid(1,1)) : End Sub : End Module"#,
    ["10"]
);
vb_full_spec!(
    array_spec_jagged_array_can_store_inner_arrays,
    r#"Module M : Sub Main() : Dim outer()() As Integer = { New Integer() {1,2}, New Integer() {3,4} } : Console.WriteLine(outer(1)(0)) : End Sub : End Module"#,
    ["3"]
);
vb_full_spec!(
    array_spec_array_of_objects_preserves_field_access,
    r#"Class C : Public Name As String : Public Sub New(v As String) : Name=v : End Sub : End Class : Module M : Sub Main() : Dim items() As C = {New C("a"), New C("b")} : Console.WriteLine(items(1).Name) : End Sub : End Module"#,
    ["b"]
);
vb_full_spec!(
    array_spec_foreach_over_string_array_preserves_order,
    r#"Module M : Sub Main() : Dim items() As String = {"a","b","c"} : Dim s As String = "" : For Each item In items : s &= item : Next : Console.WriteLine(s) : End Sub : End Module"#,
    ["abc"]
);
vb_full_spec!(
    array_spec_for_index_loop_can_write_each_array_slot,
    r#"Module M : Sub Main() : Dim values(2) As Integer : For i As Integer = 0 To 2 : values(i)=i+1 : Next : Console.WriteLine(values(2)) : End Sub : End Module"#,
    ["3"]
);
vb_full_spec!(
    array_spec_array_literal_of_booleans_can_be_counted,
    r#"Module M : Sub Main() : Dim flags() As Boolean = {True, False, True} : Dim count As Integer = 0 : For Each flag In flags : If flag Then count += 1 : Next : Console.WriteLine(count) : End Sub : End Module"#,
    ["2"]
);
vb_full_spec!(
    array_spec_array_function_can_mix_integer_values,
    r#"Module M : Sub Main() : Dim items = Array(10,20,30) : Console.WriteLine(items(1)) : End Sub : End Module"#,
    ["20"]
);
vb_full_spec!(
    array_spec_array_function_can_mix_string_values_as_objects,
    r#"Module M : Sub Main() : Dim items = Array("one","two") : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["one"]
);
vb_full_spec!(
    array_spec_lbound_and_ubound_can_drive_loop_bounds,
    r#"Module M : Sub Main() : Dim items() As Integer = {2,4,6} : Dim total As Integer = 0 : For i As Integer = LBound(items) To UBound(items) : total += items(i) : Next : Console.WriteLine(total) : End Sub : End Module"#,
    ["12"]
);
vb_full_spec!(
    array_spec_array_assignment_copies_reference_semantics,
    r#"Module M : Sub Main() : Dim a() As Integer = {1,2} : Dim b() As Integer = a : b(0)=9 : Console.WriteLine(a(0)) : End Sub : End Module"#,
    ["9"]
);
vb_full_spec!(
    array_spec_byref_parameter_can_update_array_slot,
    r#"Module M : Sub Bump(ByRef value As Integer) : value += 1 : End Sub : Sub Main() : Dim values() As Integer = {4} : Bump(values(0)) : Console.WriteLine(values(0)) : End Sub : End Module"#,
    ["5"]
);
vb_full_spec!(
    array_spec_array_of_structures_preserves_member_reads,
    r#"Structure Point : Public X As Integer : End Structure : Module M : Sub Main() : Dim points() As Point = { New Point With {.X = 2}, New Point With {.X = 7} } : Console.WriteLine(points(1).X) : End Sub : End Module"#,
    ["7"]
);
vb_full_spec!(
    array_spec_array_of_enums_can_be_compared_in_loop,
    r#"Enum Tone : Low : High : End Enum : Module M : Sub Main() : Dim tones() As Tone = {Tone.Low, Tone.High} : Dim count As Integer = 0 : For Each t In tones : If t = Tone.High Then count += 1 : Next : Console.WriteLine(count) : End Sub : End Module"#,
    ["1"]
);
vb_full_spec!(
    array_spec_nested_array_literals_can_model_matrix_rows,
    r#"Module M : Sub Main() : Dim rows()() As Integer = { New Integer(){1,2}, New Integer(){3,4} } : Console.WriteLine(rows(0)(1)+rows(1)(0)) : End Sub : End Module"#,
    ["5"]
);
vb_full_spec!(
    array_spec_preserve_can_extend_and_keep_last_existing_value,
    r#"Module M : Sub Main() : Dim values() As Integer = {1,5} : ReDim Preserve values(3) : Console.WriteLine(values(1)) : End Sub : End Module"#,
    ["5"]
);
vb_full_spec!(
    array_spec_erase_can_clear_object_array_to_nothing,
    r#"Class C : End Class : Module M : Sub Main() : Dim items() As C = {New C()} : Erase items : Console.WriteLine(IsNothing(items)) : End Sub : End Module"#,
    ["true"]
);
vb_full_spec!(
    array_spec_fixed_array_can_be_initialized_element_by_element,
    r#"Module M : Sub Main() : Dim values(1) As String : values(0)="A" : values(1)="B" : Console.WriteLine(values(0) & values(1)) : End Sub : End Module"#,
    ["AB"]
);
vb_full_spec!(
    array_spec_array_index_expression_can_use_function_result,
    r#"Module M : Function Pick() As Integer : Return 1 : End Function : Sub Main() : Dim values() As String = {"zero", "one"} : Console.WriteLine(values(Pick())) : End Sub : End Module"#,
    ["one"]
);
vb_full_spec!(
    array_spec_array_can_hold_return_values_from_helper_function,
    r#"Module M : Function NextValue() As Integer : Return 7 : End Function : Sub Main() : Dim values() As Integer = {NextValue(), NextValue()+1} : Console.WriteLine(values(1)) : End Sub : End Module"#,
    ["8"]
);
vb_full_spec!(
    array_spec_array_can_store_strings_with_embedded_spaces,
    r#"Module M : Sub Main() : Dim items() As String = {"hello world", "vb"} : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["hello world"]
);
vb_full_spec!(
    array_spec_array_iteration_can_build_delimited_string,
    r#"Module M : Sub Main() : Dim items() As String = {"a","b","c"} : Dim s As String = "" : For Each item In items : If s <> "" Then s &= "|" : s &= item : Next : Console.WriteLine(s) : End Sub : End Module"#,
    ["a|b|c"]
);
vb_full_spec!(
    array_spec_array_iteration_can_skip_nothing_entries,
    r#"Module M : Sub Main() : Dim items() As String = {"a", Nothing, "c"} : Dim s As String = "" : For Each item In items : If Not IsNothing(item) Then s &= item : Next : Console.WriteLine(s) : End Sub : End Module"#,
    ["ac"]
);
vb_full_spec!(
    array_spec_array_of_dates_can_be_printed_after_assignment,
    r#"Module M : Sub Main() : Dim items(1) As Date : items(0)=#5/14/2024# : items(1)=#5/15/2024# : Console.WriteLine(CStr(items(1))) : End Sub : End Module"#,
    ["5/15/2024"]
);
vb_full_spec!(
    array_spec_array_can_be_passed_to_sub_byref,
    r#"Module M : Sub Fill(ByRef items() As Integer) : items(0)=9 : End Sub : Sub Main() : Dim items() As Integer = {1,2} : Fill(items) : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["9"]
);
vb_full_spec!(
    array_spec_array_can_be_passed_to_function_and_return_count,
    r#"Module M : Function Count(items() As Integer) As Integer : Return UBound(items) + 1 : End Function : Sub Main() : Dim items() As Integer = {1,2,3} : Console.WriteLine(Count(items)) : End Sub : End Module"#,
    ["3"]
);
vb_full_spec!(
    array_spec_redim_preserve_can_be_used_multiple_times,
    r#"Module M : Sub Main() : Dim items() As Integer = {1} : ReDim Preserve items(1) : ReDim Preserve items(2) : items(2)=5 : Console.WriteLine(items(0)) : Console.WriteLine(items(2)) : End Sub : End Module"#,
    ["1", "5"]
);
vb_full_spec!(
    array_spec_array_item_assignment_can_use_compound_expression,
    r#"Module M : Sub Main() : Dim items() As Integer = {1,2} : items(1)=items(0)+items(1) : Console.WriteLine(items(1)) : End Sub : End Module"#,
    ["3"]
);
vb_full_spec!(
    array_spec_array_length_logic_can_use_ubound_plus_one,
    r#"Module M : Sub Main() : Dim items() As Integer = {1,2,3,4} : Console.WriteLine(UBound(items)+1) : End Sub : End Module"#,
    ["4"]
);
vb_full_spec!(
    array_spec_array_can_be_returned_from_function,
    r#"Module M : Function Build() As Integer() : Return New Integer() {4,5} : End Function : Sub Main() : Console.WriteLine(Build()(1)) : End Sub : End Module"#,
    ["5"]
);
vb_full_spec!(
    array_spec_array_can_be_received_from_function_into_local,
    r#"Module M : Function Build() As String() : Return New String() {"x","y"} : End Function : Sub Main() : Dim items() As String = Build() : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["x"]
);
vb_full_spec!(
    array_spec_array_of_lists_can_store_collection_objects,
    r#"Module M : Sub Main() : Dim bags() As List(Of Integer) = { New List(Of Integer)(), New List(Of Integer)() } : bags(1).Add(9) : Console.WriteLine(bags(1).Count) : End Sub : End Module"#,
    ["1"]
);
vb_full_spec!(
    array_spec_array_of_dictionaries_can_store_lookup_objects,
    r#"Module M : Sub Main() : Dim maps() As Dictionary(Of String, Integer) = { New Dictionary(Of String, Integer)() } : maps(0).Add("x", 7) : Console.WriteLine(maps(0).Item("x")) : End Sub : End Module"#,
    ["7"]
);
vb_full_spec!(
    array_spec_array_foreach_can_mutate_running_total,
    r#"Module M : Sub Main() : Dim items() As Integer = {1,2,3} : Dim total As Integer = 0 : For Each item In items : total += item : Next : Console.WriteLine(total) : End Sub : End Module"#,
    ["6"]
);
vb_full_spec!(
    array_spec_array_bounds_can_drive_descending_loop,
    r#"Module M : Sub Main() : Dim items() As Integer = {1,2,3} : Dim total As Integer = 0 : For i As Integer = UBound(items) To LBound(items) Step -1 : total += items(i) : Next : Console.WriteLine(total) : End Sub : End Module"#,
    ["6"]
);
vb_full_spec!(
    array_spec_list_of_integer_can_sort_values,
    r#"Module M : Sub Main() : Dim items As New List(Of Integer) : items.Add(3) : items.Add(1) : items.Add(2) : items.Sort() : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["1"]
);
vb_full_spec!(
    array_spec_list_of_string_can_reverse_order,
    r#"Module M : Sub Main() : Dim items As New List(Of String) : items.Add("a") : items.Add("b") : items.Reverse() : Console.WriteLine(items(0)) : End Sub : End Module"#,
    ["b"]
);
vb_full_spec!(
    array_spec_dictionary_can_iterate_keys_with_foreach,
    r#"Module M : Sub Main() : Dim d As New Dictionary(Of String, Integer) : d.Add("a",1) : d.Add("b",2) : Dim total As Integer = 0 : For Each key In d.Keys : total += d.Item(key) : Next : Console.WriteLine(total) : End Sub : End Module"#,
    ["3"]
);
vb_full_spec!(
    array_spec_queue_and_stack_can_model_fifo_and_lifo,
    r#"Module M : Sub Main() : Dim q As New Queue(Of Integer) : q.Enqueue(1) : q.Enqueue(2) : Dim s As New Stack(Of Integer) : s.Push(1) : s.Push(2) : Console.WriteLine(q.Dequeue()) : Console.WriteLine(s.Pop()) : End Sub : End Module"#,
    ["1", "2"]
);
