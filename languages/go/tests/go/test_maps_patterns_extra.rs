use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    map_make_set_len_runtime => ("package main; import \"fmt\"; func main() { values := make(map[string]int); values[\"a\"] = 3; fmt.Println(len(values)); }", vec!["1"]),
    map_literal_lookup_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 5}; fmt.Println(values[\"a\"]); }", vec!["5"]),
    map_comma_ok_present_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 2}; number, ok := values[\"a\"]; fmt.Println(number); fmt.Println(ok); }", vec!["2", "true"]),
    map_comma_ok_missing_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{}; number, ok := values[\"a\"]; fmt.Println(number); fmt.Println(ok); }", vec!["0", "false"]),
    map_delete_present_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; delete(values, \"a\"); fmt.Println(len(values)); }", vec!["1"]),
    map_delete_missing_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; delete(values, \"missing\"); fmt.Println(len(values)); }", vec!["1"]),
    map_update_existing_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; values[\"a\"] = 7; fmt.Println(values[\"a\"]); }", vec!["7"]),
    map_zero_value_missing_string_runtime => ("package main; import \"fmt\"; func main() { values := map[string]string{}; fmt.Println(values[\"missing\"] == \"\"); }", vec!["true"]),
    map_zero_value_missing_bool_runtime => ("package main; import \"fmt\"; func main() { values := map[string]bool{}; fmt.Println(values[\"missing\"]); }", vec!["false"]),
    nested_map_lookup_runtime => ("package main; import \"fmt\"; func main() { values := map[string]map[string]int{\"outer\": {\"inner\": 8}}; fmt.Println(values[\"outer\"][\"inner\"]); }", vec!["8"]),
    map_of_slices_len_runtime => ("package main; import \"fmt\"; func main() { values := map[string][]int{\"a\": []int{1, 2, 3}}; fmt.Println(len(values[\"a\"])); }", vec!["3"]),
    map_pass_to_function_runtime => ("package main; import \"fmt\"; func total(values map[string]int) int { return values[\"a\"] + values[\"b\"] }; func main() { fmt.Println(total(map[string]int{\"a\": 2, \"b\": 4})); }", vec!["6"]),
    map_return_from_function_runtime => ("package main; import \"fmt\"; func build() map[string]int { return map[string]int{\"a\": 6} }; func main() { fmt.Println(build()[\"a\"]); }", vec!["6"]),
    map_alias_mutation_visible_runtime => ("package main; import \"fmt\"; func main() { left := map[string]int{\"a\": 1}; right := left; right[\"a\"] = 9; fmt.Println(left[\"a\"]); }", vec!["9"]),
    map_in_struct_field_runtime => ("package main; import \"fmt\"; type holder struct { values map[string]int }; func main() { value := holder{values: map[string]int{\"a\": 4}}; fmt.Println(value.values[\"a\"]); }", vec!["4"]),
    map_int_keys_sum_runtime => ("package main; import \"fmt\"; func main() { values := map[int]int{1: 2, 2: 3}; fmt.Println(values[1] + values[2]); }", vec!["5"]),
    map_bool_keys_runtime => ("package main; import \"fmt\"; func main() { values := map[bool]string{true: \"yes\", false: \"no\"}; fmt.Println(values[true]); }", vec!["yes"]),
    map_string_concat_values_runtime => ("package main; import \"fmt\"; func main() { values := map[string]string{\"a\": \"vy\", \"b\": \"be\"}; fmt.Println(values[\"a\"] + values[\"b\"]); }", vec!["vybe"]),
    map_increment_pattern_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 2}; values[\"a\"] = values[\"a\"] + 5; fmt.Println(values[\"a\"]); }", vec!["7"]),
    nil_map_lookup_len_runtime => ("package main; import \"fmt\"; func main() { var values map[string]int; fmt.Println(values[\"a\"]); fmt.Println(len(values)); }", vec!["0", "0"]),
    map_range_sum_values_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 2, \"b\": 5}; total := 0; for _, value := range values { total += value }; fmt.Println(total); }", vec!["7"]),
    map_range_count_entries_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2, \"c\": 3}; count := 0; for range values { count++ }; fmt.Println(count); }", vec!["3"]),
    map_of_struct_value_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { values := map[string]point{\"a\": {x: 13}}; fmt.Println(values[\"a\"].x); }", vec!["13"]),
    map_of_arrays_value_runtime => ("package main; import \"fmt\"; func main() { values := map[string][2]int{\"a\": [2]int{3, 4}}; fmt.Println(values[\"a\"][1]); }", vec!["4"]),
    delete_then_lookup_ok_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; delete(values, \"a\"); _, ok := values[\"a\"]; fmt.Println(ok); }", vec!["false"]),
}

compile_cases! {
    map_with_struct_key_compile => "package main; type point struct { x int; y int }; func main() { _ = map[point]int{} }",
    map_with_array_key_compile => "package main; func main() { _ = map[[2]int]string{} }",
    nested_map_make_compile => "package main; func main() { values := make(map[string]map[string]int); _ = values }",
    map_range_blank_identifiers_compile => "package main; func main() { values := map[string]int{}; for _, value := range values { _ = value } }",
    nil_map_delete_compile => "package main; func main() { var values map[string]int; delete(values, \"a\") }",
    map_of_slices_compile => "package main; func main() { _ = map[string][]int{\"a\": []int{1}} }",
    map_of_maps_compile => "package main; func main() { _ = map[string]map[string]int{\"a\": map[string]int{\"b\": 1}} }",
    map_of_structs_compile => "package main; type point struct { x int }; func main() { _ = map[string]point{\"a\": {x: 1}} }",
    map_assignment_compile => "package main; func main() { values := map[string]int{}; values[\"a\"] = 1; _ = values }",
    map_lookup_blank_identifier_compile => "package main; func main() { values := map[string]int{\"a\": 1}; _, _ = values[\"a\"] }",
    map_make_with_len_compile => "package main; func main() { values := make(map[string]int, 4); _ = values }",
    map_delete_after_lookup_compile => "package main; func main() { values := map[string]int{\"a\": 1}; _ = values[\"a\"]; delete(values, \"a\") }",
    map_bool_key_compile => "package main; func main() { _ = map[bool]int{true: 1} }",
    map_int_key_compile => "package main; func main() { _ = map[int]string{1: \"one\"} }",
    map_string_key_compile => "package main; func main() { _ = map[string]int{\"a\": 1} }",
    map_value_struct_field_compile => "package main; type point struct { x int }; func main() { values := map[string]point{\"a\": {x: 1}}; _ = values[\"a\"].x }",
    map_value_array_index_compile => "package main; func main() { values := map[string][2]int{\"a\": [2]int{1, 2}}; _ = values[\"a\"][1] }",
    map_in_struct_compile => "package main; type holder struct { values map[string]int }; func main() { _ = holder{} }",
    map_return_compile => "package main; func build() map[string]int { return map[string]int{} }; func main() { _ = build() }",
    map_parameter_compile => "package main; func use(values map[string]int) int { return len(values) }; func main() { _ = use }",
    map_compare_to_nil_compile => "package main; func main() { var values map[string]int; _ = (values == nil) }",
    map_lookup_missing_compile => "package main; func main() { values := map[string]int{}; _ = values[\"missing\"] }",
    map_value_increment_compile => "package main; func main() { values := map[string]int{\"a\": 1}; values[\"a\"] = values[\"a\"] + 1 }",
    map_range_keys_values_compile => "package main; func main() { values := map[string]int{\"a\": 1}; for key, value := range values { _, _ = key, value } }",
    map_literal_with_trailing_comma_compile => "package main; func main() { _ = map[string]int{\"a\": 1, \"b\": 2,} }",
}
