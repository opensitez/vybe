use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    array_index_of_finds_first_matching_element,
    r#"var values = new[] { 4, 7, 9 }; Console.WriteLine(System.Array.IndexOf(values, 7));"#,
    ["1"]
);
csharp_case!(
    array_last_index_of_finds_last_matching_element,
    r#"var values = new[] { 1, 2, 1, 3 }; Console.WriteLine(System.Array.LastIndexOf(values, 1));"#,
    ["2"]
);
csharp_case!(
    array_reverse_reorders_items_in_place,
    r#"var values = new[] { 1, 2, 3 }; System.Array.Reverse(values); foreach (var value in values) Console.WriteLine(value);"#,
    ["3", "2", "1"]
);
csharp_case!(
    array_clear_resets_values_to_default,
    r#"var values = new[] { 1, 2, 3 }; System.Array.Clear(values, 1, 2); foreach (var value in values) Console.WriteLine(value);"#,
    ["1", "0", "0"]
);
csharp_case!(
    array_copy_moves_values_between_arrays,
    r#"var source = new[] { 5, 6, 7 }; var target = new int[3]; System.Array.Copy(source, target, 3); foreach (var value in target) Console.WriteLine(value);"#,
    ["5", "6", "7"]
);
csharp_case!(
    array_resize_grows_array_and_preserves_existing_values,
    r#"var values = new[] { 2, 4 }; System.Array.Resize(ref values, 4); foreach (var value in values) Console.WriteLine(value);"#,
    ["2", "4", "0", "0"]
);
csharp_case!(
    array_sort_orders_values_ascending,
    r#"var values = new[] { 4, 1, 3 }; System.Array.Sort(values); foreach (var value in values) Console.WriteLine(value);"#,
    ["1", "3", "4"]
);
csharp_case!(
    array_binary_search_returns_position_of_found_value,
    r#"var values = new[] { 1, 3, 5, 7 }; Console.WriteLine(System.Array.BinarySearch(values, 5));"#,
    ["2"]
);
csharp_case!(
    array_exists_reports_true_when_predicate_matches,
    r#"var values = new[] { 1, 3, 5 }; Console.WriteLine(System.Array.Exists(values, value => value == 3));"#,
    ["True"]
);
csharp_case!(
    array_find_returns_first_matching_value,
    r#"var values = new[] { 2, 4, 5, 8 }; Console.WriteLine(System.Array.Find(values, value => value % 2 == 1));"#,
    ["5"]
);
csharp_case!(
    array_find_index_returns_position_of_match,
    r#"var values = new[] { 2, 4, 5, 8 }; Console.WriteLine(System.Array.FindIndex(values, value => value % 2 == 1));"#,
    ["2"]
);
csharp_case!(
    array_convert_all_maps_values_to_new_type,
    r#"var values = new[] { 1, 2, 3 }; var text = System.Array.ConvertAll(values, value => "n" + value); foreach (var value in text) Console.WriteLine(value);"#,
    ["n1", "n2", "n3"]
);
csharp_case!(
    array_true_for_all_checks_entire_sequence,
    r#"var values = new[] { 2, 4, 6 }; Console.WriteLine(System.Array.TrueForAll(values, value => value % 2 == 0));"#,
    ["True"]
);
csharp_case!(
    array_empty_returns_zero_length_array,
    r#"Console.WriteLine(System.Array.Empty<string>().Length);"#,
    ["0"]
);
csharp_case!(
    array_create_instance_builds_runtime_sized_array,
    r#"var array = System.Array.CreateInstance(typeof(int), 3); Console.WriteLine(array.Length);"#,
    ["3"]
);
csharp_case!(
    array_rank_reports_dimension_count_for_vector,
    r#"var values = new[] { 1, 2, 3 }; Console.WriteLine(values.Rank);"#,
    ["1"]
);
csharp_case!(
    multidimensional_array_get_length_reports_dimension_size,
    r#"var grid = new int[2, 3]; Console.WriteLine(grid.GetLength(0)); Console.WriteLine(grid.GetLength(1));"#,
    ["2", "3"]
);
csharp_case!(
    instance_copy_to_moves_values_into_target_array,
    r#"var source = new[] { 9, 8 }; var target = new int[2]; source.CopyTo(target, 0); foreach (var value in target) Console.WriteLine(value);"#,
    ["9", "8"]
);
csharp_case!(
    array_clone_creates_independent_shallow_copy,
    r#"var source = new[] { 1, 2 }; var clone = (int[])source.Clone(); clone[0] = 9; Console.WriteLine(source[0]); Console.WriteLine(clone[0]);"#,
    ["1", "9"]
);
csharp_case!(
    array_for_each_invokes_action_for_each_item,
    r#"var values = new[] { 3, 4 }; System.Array.ForEach(values, value => Console.WriteLine(value * 2));"#,
    ["6", "8"]
);
