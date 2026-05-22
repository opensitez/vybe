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
    subslice_reads_selected_window_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; part := values[1:]; fmt.Println(part[0]); fmt.Println(len(part)); }", vec!["2", "2"]),
    prefix_slice_len_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[:2]; fmt.Println(len(part)); }", vec!["2"]),
    suffix_slice_len_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[2:]; fmt.Println(len(part)); }", vec!["2"]),
    mid_slice_read_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[1:3]; fmt.Println(part[1]); }", vec!["3"]),
    reslice_growth_within_capacity_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[:2]; part = part[:4]; fmt.Println(part[3]); }", vec!["4"]),
    copy_slice_returns_count_runtime => ("package main; import \"fmt\"; func main() { dst := make([]int, 2); src := []int{3, 4, 5}; fmt.Println(copy(dst, src)); }", vec!["2"]),
    copy_slice_copies_values_runtime => ("package main; import \"fmt\"; func main() { dst := make([]int, 2); src := []int{3, 4}; copy(dst, src); fmt.Println(dst[1]); }", vec!["4"]),
    append_slice_expansion_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2}; values = append(values, 3, 4); fmt.Println(len(values)); fmt.Println(values[3]); }", vec!["4", "4"]),
    subslice_observes_source_mutation_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; part := values[:2]; values[0] = 7; fmt.Println(part[0]); }", vec!["7"]),
    slice_literal_zero_len_runtime => ("package main; import \"fmt\"; func main() { values := []int{}; fmt.Println(len(values)); }", vec!["0"]),
    make_slice_len_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 3); fmt.Println(len(values)); }", vec!["3"]),
    make_slice_len_cap_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 2, 5); fmt.Println(len(values)); fmt.Println(cap(values)); }", vec!["2", "5"]),
    append_to_make_slice_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 0, 3); values = append(values, 6); fmt.Println(values[0]); }", vec!["6"]),
    slice_of_structs_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { values := []point{{x: 8}, {x: 9}}; fmt.Println(values[1].x); }", vec!["9"]),
    slice_of_strings_index_runtime => ("package main; import \"fmt\"; func main() { values := []string{\"vy\", \"be\"}; fmt.Println(values[0] + values[1]); }", vec!["vybe"]),
    nested_slice_lookup_runtime => ("package main; import \"fmt\"; func main() { values := [][]int{{1, 2}, {3, 4}}; fmt.Println(values[1][0]); }", vec!["3"]),
    slice_pass_to_function_runtime => ("package main; import \"fmt\"; func total(values []int) int { return values[0] + values[1] }; func main() { fmt.Println(total([]int{4, 5})); }", vec!["9"]),
    slice_return_from_function_runtime => ("package main; import \"fmt\"; func build() []int { return []int{5, 6} }; func main() { fmt.Println(build()[1]); }", vec!["6"]),
    slice_swap_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2}; values[0], values[1] = values[1], values[0]; fmt.Println(values[0]); fmt.Println(values[1]); }", vec!["2", "1"]),
    slice_index_assignment_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; values[2] = 7; fmt.Println(values[2]); }", vec!["7"]),
    slice_zero_value_from_make_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 2); fmt.Println(values[0]); fmt.Println(values[1]); }", vec!["0", "0"]),
    slice_alias_after_assignment_runtime => ("package main; import \"fmt\"; func main() { left := []int{1, 2}; right := left; right[0] = 8; fmt.Println(left[0]); }", vec!["8"]),
    slice_append_preserves_prefix_runtime => ("package main; import \"fmt\"; func main() { values := []int{2, 4}; values = append(values, 6); fmt.Println(values[0]); fmt.Println(values[2]); }", vec!["2", "6"]),
    slice_len_after_copy_runtime => ("package main; import \"fmt\"; func main() { dst := make([]int, 3); copy(dst, []int{1, 2}); fmt.Println(len(dst)); }", vec!["3"]),
    slice_subslice_of_subslice_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[1:4]; next := part[1:]; fmt.Println(next[1]); }", vec!["4"]),
}

compile_cases! {
    subslice_shares_backing_array_compile => "package main; func main() { values := []int{1, 2, 3}; part := values[1:]; part[0] = 9; _ = values[1] }",
    full_slice_expression_compile => "package main; func main() { values := []int{1, 2, 3}; _ = values[0:2:3] }",
    slice_from_array_compile => "package main; func main() { values := [3]int{1, 2, 3}; _ = values[1:] }",
    slice_to_end_compile => "package main; func main() { values := []int{1, 2, 3}; _ = values[1:] }",
    slice_from_start_compile => "package main; func main() { values := []int{1, 2, 3}; _ = values[:2] }",
    slice_mid_compile => "package main; func main() { values := []int{1, 2, 3}; _ = values[1:2] }",
    nested_slice_literal_compile => "package main; func main() { _ = [][]int{{1, 2}, {3, 4}} }",
    append_slice_spread_compile => "package main; func main() { values := []int{1}; extra := []int{2, 3}; values = append(values, extra...); _ = values }",
    copy_into_nil_slice_compile => "package main; func main() { var dst []int; src := []int{1}; _ = copy(dst, src) }",
    copy_from_nil_slice_compile => "package main; func main() { dst := make([]int, 1); var src []int; _ = copy(dst, src) }",
    make_slice_with_cap_compile => "package main; func main() { _ = make([]int, 2, 4) }",
    make_nested_slice_compile => "package main; func main() { values := make([][]int, 2); _ = values }",
    slice_field_in_struct_compile => "package main; type holder struct { values []int }; func main() { _ = holder{} }",
    slicing_in_return_compile => "package main; func build(values []int) []int { return values[1:] }; func main() { _ = build }",
    reslice_nil_compile => "package main; func main() { var values []int; _ = values[:] }",
    slice_assignment_compile => "package main; func main() { values := []int{1, 2}; values[0] = 3 }",
    slice_swap_assignment_compile => "package main; func main() { values := []int{1, 2}; values[0], values[1] = values[1], values[0] }",
    append_into_struct_field_compile => "package main; type holder struct { values []int }; func main() { value := holder{}; value.values = append(value.values, 1); _ = value }",
    nested_slice_index_compile => "package main; func main() { values := [][]int{{1}}; _ = values[0][0] }",
    slice_pass_to_function_compile => "package main; func use(values []int) int { return len(values) }; func main() { _ = use([]int{1}) }",
    slice_return_compile => "package main; func build() []int { return []int{1, 2} }; func main() { _ = build() }",
    slice_alias_compile => "package main; func main() { left := []int{1}; right := left; _ = right }",
    append_nil_slice_spread_compile => "package main; func main() { var values []int; extra := []int{1, 2}; values = append(values, extra...); _ = values }",
    copy_result_used_compile => "package main; func main() { dst := make([]int, 2); src := []int{1}; n := copy(dst, src); _ = n }",
    slice_of_structs_compile => "package main; type point struct { x int }; func main() { _ = []point{{x: 1}} }",
    slice_of_maps_compile => "package main; func main() { _ = []map[string]int{{\"a\": 1}} }",
}