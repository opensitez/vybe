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
    range_over_slice_sum_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range []int{1, 2, 3} { total += value }; fmt.Println(total); }", vec!["6"]),
    range_over_array_sum_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range [3]int{2, 3, 4} { total += value }; fmt.Println(total); }", vec!["9"]),
    range_over_string_runes_sum_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range \"AB\" { total += int(value) }; fmt.Println(total); }", vec!["131"]),
    range_over_slice_of_strings_count_runtime => ("package main; import \"fmt\"; func main() { count := 0; for _, value := range []string{\"go\", \"vybe\"} { count += len(value) }; fmt.Println(count); }", vec!["6"]),
    range_over_map_sum_values_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 2, \"b\": 5}; total := 0; for _, value := range values { total += value }; fmt.Println(total); }", vec!["7"]),
    range_over_slice_index_sum_runtime => ("package main; import \"fmt\"; func main() { total := 0; for index := range []int{4, 5, 6} { total += index }; fmt.Println(total); }", vec!["3"]),
    range_over_slice_value_only_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range []int{3, 3, 3} { total += value }; fmt.Println(total); }", vec!["9"]),
    range_over_array_index_only_runtime => ("package main; import \"fmt\"; func main() { total := 0; for index := range [2]int{7, 8} { total += index }; fmt.Println(total); }", vec!["1"]),
    range_over_nil_slice_count_runtime => ("package main; import \"fmt\"; func main() { var values []int; count := 0; for range values { count++ }; fmt.Println(count); }", vec!["0"]),
    range_over_nil_map_count_runtime => ("package main; import \"fmt\"; func main() { var values map[string]int; count := 0; for range values { count++ }; fmt.Println(count); }", vec!["0"]),
    range_over_slice_break_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range []int{1, 2, 3} { total += value; break }; fmt.Println(total); }", vec!["1"]),
    range_over_slice_continue_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range []int{1, 2, 3} { if value == 2 { continue }; total += value }; fmt.Println(total); }", vec!["4"]),
    range_over_slice_blank_identifier_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, value := range []int{2, 4} { total += value }; fmt.Println(total); }", vec!["6"]),
    range_over_string_index_sum_runtime => ("package main; import \"fmt\"; func main() { total := 0; for index := range \"go\" { total += index }; fmt.Println(total); }", vec!["1"]),
    range_over_nested_slice_runtime => ("package main; import \"fmt\"; func main() { total := 0; for _, row := range [][]int{{1, 2}, {3}} { total += len(row) }; fmt.Println(total); }", vec!["3"]),
    range_over_slice_modify_by_index_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; for index := range values { values[index]++ }; fmt.Println(values[0]); fmt.Println(values[2]); }", vec!["2", "4"]),
    range_over_map_count_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; count := 0; for range values { count++ }; fmt.Println(count); }", vec!["2"]),
    range_over_slice_build_sum_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 3); values[0], values[1], values[2] = 1, 2, 3; total := 0; for _, value := range values { total += value }; fmt.Println(total); }", vec!["6"]),
    range_over_string_count_runtime => ("package main; import \"fmt\"; func main() { count := 0; for range \"vybe\" { count++ }; fmt.Println(count); }", vec!["4"]),
    range_with_outer_accumulator_runtime => ("package main; import \"fmt\"; func main() { total := 1; for _, value := range []int{2, 3} { total *= value }; fmt.Println(total); }", vec!["6"]),
    range_over_struct_slice_field_runtime => ("package main; import \"fmt\"; type holder struct { values []int }; func main() { value := holder{values: []int{1, 2, 3}}; total := 0; for _, item := range value.values { total += item }; fmt.Println(total); }", vec!["6"]),
    range_over_map_of_slices_runtime => ("package main; import \"fmt\"; func main() { values := map[string][]int{\"a\": []int{1, 2}, \"b\": []int{3}}; total := 0; for _, item := range values { total += len(item) }; fmt.Println(total); }", vec!["3"]),
    range_over_array_values_runtime => ("package main; import \"fmt\"; func main() { last := 0; for _, value := range [3]int{4, 5, 6} { last = value }; fmt.Println(last); }", vec!["6"]),
    range_over_empty_slice_runtime => ("package main; import \"fmt\"; func main() { count := 0; for range []int{} { count++ }; fmt.Println(count); }", vec!["0"]),
    range_over_make_slice_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 2); values[0], values[1] = 3, 4; total := 0; for _, value := range values { total += value }; fmt.Println(total); }", vec!["7"]),
    range_over_string_last_rune_runtime => ("package main; import \"fmt\"; func main() { last := rune(0); for _, value := range \"go\" { last = value }; fmt.Println(int(last)); }", vec!["111"]),
}

compile_cases! {
    range_over_channel_compile => "package main; func main() { ch := make(chan int); for value := range ch { _ = value } }",
    range_over_string_blank_identifier_compile => "package main; func main() { for _, value := range \"go\" { _ = value } }",
    range_over_map_key_value_compile => "package main; func main() { values := map[string]int{\"a\": 1}; for key, value := range values { _, _ = key, value } }",
    nested_range_compile => "package main; func main() { for _, row := range [][]int{{1}} { for _, value := range row { _ = value } } }",
    labeled_break_in_range_compile => "package main; func main() { outer: for _, row := range [][]int{{1}} { for _, value := range row { _ = value; break outer } } }",
    labeled_continue_in_range_compile => "package main; func main() { outer: for _, value := range []int{1, 2} { if value == 1 { continue outer } } }",
    range_assign_existing_vars_compile => "package main; func main() { values := []int{1}; var index int; var value int; for index, value = range values { _, _ = index, value } }",
    range_index_only_compile => "package main; func main() { values := []int{1}; for index := range values { _ = index } }",
    range_value_only_compile => "package main; func main() { values := []int{1}; for _, value := range values { _ = value } }",
    range_over_array_literal_compile => "package main; func main() { for _, value := range [2]int{1, 2} { _ = value } }",
    range_over_slice_literal_compile => "package main; func main() { for _, value := range []int{1, 2} { _ = value } }",
    range_over_map_literal_compile => "package main; func main() { for key, value := range map[string]int{\"a\": 1} { _, _ = key, value } }",
    range_over_nil_slice_compile => "package main; func main() { var values []int; for _, value := range values { _ = value } }",
    range_over_nil_map_compile => "package main; func main() { var values map[string]int; for key := range values { _ = key } }",
    range_over_struct_field_compile => "package main; type holder struct { values []int }; func main() { value := holder{values: []int{1}}; for _, item := range value.values { _ = item } }",
    range_with_if_compile => "package main; func main() { for _, value := range []int{1, 2} { if value > 0 { _ = value } } }",
    range_with_switch_compile => "package main; func main() { for _, value := range []int{1} { switch value { case 1: _ = value } } }",
    range_with_break_compile => "package main; func main() { for _, value := range []int{1} { _ = value; break } }",
    range_with_continue_compile => "package main; func main() { for _, value := range []int{1, 2} { if value == 1 { continue }; _ = value } }",
    range_with_short_decl_body_compile => "package main; func main() { for _, value := range []int{1} { next := value + 1; _ = next } }",
    range_over_string_compile => "package main; func main() { for index, value := range \"go\" { _, _ = index, value } }",
    range_over_make_slice_compile => "package main; func main() { values := make([]int, 2); for _, value := range values { _ = value } }",
    range_over_nested_map_compile => "package main; func main() { values := map[string]map[string]int{\"a\": {\"b\": 1}}; for _, item := range values { _ = item } }",
    range_over_map_of_slices_compile => "package main; func main() { values := map[string][]int{\"a\": []int{1}}; for _, item := range values { _ = item } }",
    range_inside_function_compile => "package main; func use(values []int) int { total := 0; for _, value := range values { total += value }; return total }; func main() { _ = use }",
}
