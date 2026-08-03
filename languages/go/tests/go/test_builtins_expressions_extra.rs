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
    len_on_array_runtime => ("package main; import \"fmt\"; func main() { values := [4]int{1, 2, 3, 4}; fmt.Println(len(values)); }", vec!["4"]),
    len_on_slice_literal_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; fmt.Println(len(values)); }", vec!["3"]),
    len_on_map_literal_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; fmt.Println(len(values)); }", vec!["2"]),
    cap_on_array_slice_runtime => ("package main; import \"fmt\"; func main() { values := [5]int{1, 2, 3, 4, 5}; part := values[1:3]; fmt.Println(cap(part)); }", vec!["4"]),
    cap_on_make_slice_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 2, 5); fmt.Println(cap(values)); }", vec!["5"]),
    append_multiple_values_runtime => ("package main; import \"fmt\"; func main() { values := []int{1}; values = append(values, 2, 3, 4); fmt.Println(len(values)); fmt.Println(values[3]); }", vec!["4", "4"]),
    copy_between_slices_runtime => ("package main; import \"fmt\"; func main() { dst := []int{0, 0, 0}; src := []int{7, 8}; copy(dst, src); fmt.Println(dst[0]); fmt.Println(dst[1]); }", vec!["7", "8"]),
    delete_existing_map_key_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; delete(values, \"a\"); fmt.Println(len(values)); }", vec!["1"]),
    delete_missing_map_key_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; delete(values, \"missing\"); fmt.Println(len(values)); }", vec!["1"]),
    make_slice_with_length_and_capacity_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 3, 6); fmt.Println(len(values)); fmt.Println(cap(values)); }", vec!["3", "6"]),
    new_int_pointer_runtime => ("package main; import \"fmt\"; func main() { value := new(int); *value = 11; fmt.Println(*value); }", vec!["11"]),
    new_struct_pointer_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { value := new(point); value.x = 12; fmt.Println(value.x); }", vec!["12"]),
    append_named_slice_runtime => ("package main; import \"fmt\"; type numbers []int; func main() { values := numbers{1, 2}; values = append(values, 3); fmt.Println(values[2]); }", vec!["3"]),
    copy_short_destination_runtime => ("package main; import \"fmt\"; func main() { dst := []int{0}; src := []int{4, 5}; fmt.Println(copy(dst, src)); fmt.Println(dst[0]); }", vec!["1", "4"]),
    copy_short_source_runtime => ("package main; import \"fmt\"; func main() { dst := []int{0, 0, 0}; src := []int{4}; fmt.Println(copy(dst, src)); fmt.Println(dst[0]); fmt.Println(dst[1]); }", vec!["1", "4", "0"]),
    len_on_nested_slice_runtime => ("package main; import \"fmt\"; func main() { grid := [][]int{{1}, {2}, {3}}; fmt.Println(len(grid)); }", vec!["3"]),
    cap_after_reslice_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3, 4}; part := values[1:2]; fmt.Println(cap(part)); }", vec!["3"]),
    make_slice_zero_length_nonzero_cap_runtime => ("package main; import \"fmt\"; func main() { values := make([]int, 0, 4); fmt.Println(len(values)); fmt.Println(cap(values)); }", vec!["0", "4"]),
    append_to_nil_slice_runtime => ("package main; import \"fmt\"; func main() { var values []int; values = append(values, 9); fmt.Println(len(values)); fmt.Println(values[0]); }", vec!["1", "9"]),
    delete_then_lookup_ok_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; delete(values, \"a\"); _, ok := values[\"a\"]; fmt.Println(ok); }", vec!["false"]),
    len_after_delete_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1, \"b\": 2}; delete(values, \"b\"); fmt.Println(len(values)); }", vec!["1"]),
    make_map_then_assign_runtime => ("package main; import \"fmt\"; func main() { values := make(map[string]int); values[\"go\"] = 7; fmt.Println(values[\"go\"]); }", vec!["7"]),
    append_result_reassigned_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2}; values = append(values, 5); fmt.Println(values[2]); }", vec!["5"]),
    copy_returns_count_runtime => ("package main; import \"fmt\"; func main() { dst := []int{0, 0}; src := []int{3, 4, 5}; fmt.Println(copy(dst, src)); }", vec!["2"]),
    cap_on_slice_literal_runtime => ("package main; import \"fmt\"; func main() { values := []int{1, 2, 3}; fmt.Println(cap(values)); }", vec!["3"]),
    append_struct_values_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { values := []point{}; values = append(values, point{x: 14}); fmt.Println(values[0].x); }", vec!["14"]),
    len_on_empty_map_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{}; fmt.Println(len(values)); }", vec!["0"]),
    len_on_pointer_dereference_runtime => ("package main; import \"fmt\"; func main() { values := &[3]int{1, 2, 3}; fmt.Println(len(*values)); }", vec!["3"]),
    cap_on_pointer_dereference_runtime => ("package main; import \"fmt\"; func main() { values := &[3]int{1, 2, 3}; fmt.Println(cap(*values)); }", vec!["3"]),
    make_bool_slice_runtime => ("package main; import \"fmt\"; func main() { flags := make([]bool, 2); fmt.Println(len(flags)); fmt.Println(flags[1]); }", vec!["2", "false"]),
    append_after_copy_runtime => ("package main; import \"fmt\"; func main() { dst := []int{0, 0}; copy(dst, []int{1, 2}); dst = append(dst, 3); fmt.Println(dst[2]); }", vec!["3"]) }

compile_cases! {
    len_on_channel_compile => "package main; func main() { ch := make(chan int, 2); _ = len(ch) }",
    append_slice_expansion_compile => "package main; func main() { a := []int{1}; b := []int{2, 3}; a = append(a, b...); _ = a }",
    append_string_bytes_compile => "package main; func main() { dst := []byte{'a'}; dst = append(dst, []byte(\"bc\")...); _ = dst }",
    make_map_with_capacity_compile => "package main; func main() { values := make(map[string]int, 4); _ = values }",
    make_channel_buffered_compile => "package main; func main() { ch := make(chan string, 3); _ = ch }",
    complex_literal_compile => "package main; func main() { z := complex(1, 2); _ = z }",
    real_builtin_compile => "package main; func main() { r := real(complex(3, 4)); _ = r }",
    imag_builtin_compile => "package main; func main() { i := imag(complex(3, 4)); _ = i }",
    close_channel_compile => "package main; func main() { ch := make(chan int); close(ch) }",
    panic_string_compile => "package main; func main() { panic(\"stop\") }",
    recover_value_compile => "package main; func main() { defer func() { _ = recover() }(); panic(\"boom\") }",
    recover_in_deferred_closure_compile => "package main; func main() { defer func(message string) { _ = recover(); _ = message }(\"done\"); panic(\"boom\") }",
    new_array_pointer_compile => "package main; func main() { values := new([3]int); _ = values }",
    make_nested_slice_compile => "package main; func main() { values := make([][]int, 2); _ = values }",
    new_named_type_pointer_compile => "package main; type counter int; func main() { value := new(counter); _ = value }",
    copy_string_into_byte_buffer_compile => "package main; func main() { dst := make([]byte, 4); _ = copy(dst, \"go\") }",
    len_on_function_returned_slice_compile => "package main; func values() []int { return []int{1, 2, 3} }; func main() { _ = len(values()) }",
    cap_on_function_returned_slice_compile => "package main; func values() []int { return make([]int, 2, 5) }; func main() { _ = cap(values()) }",
    complex_parts_in_assignment_compile => "package main; func main() { r, i := real(complex(5, 6)), imag(complex(5, 6)); _, _ = r, i }" }
