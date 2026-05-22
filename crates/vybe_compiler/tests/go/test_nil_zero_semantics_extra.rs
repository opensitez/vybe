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
    nil_slice_len_runtime => ("package main; import \"fmt\"; func main() { var values []int; fmt.Println(values == nil); fmt.Println(len(values)); }", vec!["true", "0"]),
    nil_map_read_runtime => ("package main; import \"fmt\"; func main() { var values map[string]int; fmt.Println(values == nil); fmt.Println(values[\"missing\"]); }", vec!["true", "0"]),
    nil_pointer_compare_runtime => ("package main; import \"fmt\"; func main() { var value *int; fmt.Println(value == nil); }", vec!["true"]),
    nil_interface_compare_runtime => ("package main; import \"fmt\"; func main() { var value interface{}; fmt.Println(value == nil); }", vec!["true"]),
    nil_function_compare_runtime => ("package main; import \"fmt\"; func main() { var fn func(); fmt.Println(fn == nil); }", vec!["true"]),
    nil_channel_compare_runtime => ("package main; import \"fmt\"; func main() { var ch chan int; fmt.Println(ch == nil); }", vec!["true"]),
    zero_value_struct_fields_runtime => ("package main; import \"fmt\"; type counter struct { n int; ok bool }; func main() { var c counter; fmt.Println(c.n); fmt.Println(c.ok); }", vec!["0", "false"]),
    zero_value_array_elements_runtime => ("package main; import \"fmt\"; func main() { var values [2]int; fmt.Println(values[0]); fmt.Println(values[1]); }", vec!["0", "0"]),
    zero_value_bool_runtime => ("package main; import \"fmt\"; func main() { var flag bool; fmt.Println(flag); }", vec!["false"]),
    zero_value_string_compare_runtime => ("package main; import \"fmt\"; func main() { var text string; fmt.Println(text == \"\"); }", vec!["true"]),
    zero_value_named_type_runtime => ("package main; import \"fmt\"; type score int; func main() { var s score; fmt.Println(s); }", vec!["0"]),
    nil_slice_append_runtime => ("package main; import \"fmt\"; func main() { var values []int; values = append(values, 4); fmt.Println(values[0]); }", vec!["4"]),
    nil_map_lookup_ok_runtime => ("package main; import \"fmt\"; func main() { var values map[string]int; _, ok := values[\"a\"]; fmt.Println(ok); }", vec!["false"]),
    zero_value_pointer_field_in_struct_runtime => ("package main; import \"fmt\"; type node struct { next *node }; func main() { var n node; fmt.Println(n.next == nil); }", vec!["true"]),
    zero_value_embedded_struct_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { var value outer; fmt.Println(value.count); }", vec!["0"]),
    zero_value_interface_field_runtime => ("package main; import \"fmt\"; type holder struct { value interface{} }; func main() { var h holder; fmt.Println(h.value == nil); }", vec!["true"]),
    zero_value_channel_in_struct_runtime => ("package main; import \"fmt\"; type holder struct { ch chan int }; func main() { var h holder; fmt.Println(h.ch == nil); }", vec!["true"]),
    zero_value_func_field_runtime => ("package main; import \"fmt\"; type holder struct { fn func() }; func main() { var h holder; fmt.Println(h.fn == nil); }", vec!["true"]),
    nil_map_delete_runtime => ("package main; import \"fmt\"; func main() { var values map[string]int; delete(values, \"x\"); fmt.Println(values == nil); }", vec!["true"]),
    nil_slice_copy_runtime => ("package main; import \"fmt\"; func main() { var dst []int; src := []int{1, 2}; fmt.Println(copy(dst, src)); }", vec!["0"]),
    zero_value_nested_struct_runtime => ("package main; import \"fmt\"; type inner struct { n int }; type outer struct { value inner }; func main() { var v outer; fmt.Println(v.value.n); }", vec!["0"]),
    zero_value_slice_field_len_runtime => ("package main; import \"fmt\"; type bag struct { values []int }; func main() { var b bag; fmt.Println(len(b.values)); }", vec!["0"]),
    zero_value_map_field_nil_compare_runtime => ("package main; import \"fmt\"; type bag struct { values map[string]int }; func main() { var b bag; fmt.Println(b.values == nil); }", vec!["true"]),
    zero_value_array_field_runtime => ("package main; import \"fmt\"; type bag struct { values [2]int }; func main() { var b bag; fmt.Println(b.values[1]); }", vec!["0"]),
    nil_pointer_passed_to_function_runtime => ("package main; import \"fmt\"; func isNil(v *int) bool { return v == nil }; func main() { var value *int; fmt.Println(isNil(value)); }", vec!["true"]),
}

compile_cases! {
    nil_map_assignment_compile => "package main; func main() { var values map[string]int; values[\"a\"] = 1 }",
    nil_channel_close_compile => "package main; func main() { var ch chan int; close(ch) }",
    zero_value_interface_switch_compile => "package main; func main() { var value interface{}; switch value.(type) { default: } }",
    zero_value_struct_literal_compile => "package main; type counter struct { n int }; func main() { _ = counter{} }",
    zero_value_named_pointer_compile => "package main; type score int; func main() { var value *score; _ = value }",
    nil_slice_reslice_compile => "package main; func main() { var values []int; _ = values[:] }",
    nil_map_range_compile => "package main; func main() { var values map[string]int; for k, v := range values { _, _ = k, v } }",
    nil_function_guard_compile => "package main; func main() { var fn func(); if fn == nil { return } }",
    zero_value_recursive_struct_compile => "package main; type node struct { next *node }; func main() { var n node; _ = n }",
    zero_value_channel_receive_compile => "package main; func main() { var ch chan int; _ = ch }",
    zero_value_pointer_deref_compile => "package main; func main() { var value *int; _ = *value }",
    blank_identifier_zero_values_compile => "package main; func main() { var a int; var b string; _, _ = a, b }",
    nil_interface_type_assertion_compile => "package main; func main() { var value interface{}; _, _ = value.(int) }",
    zero_value_array_compare_compile => "package main; func main() { var a [2]int; var b [2]int; _ = (a == b) }",
    nil_slice_append_multiple_compile => "package main; func main() { var values []int; values = append(values, 1, 2, 3); _ = values }",
    nil_map_len_compile => "package main; func main() { var values map[string]int; _ = len(values) }",
    nil_channel_len_compile => "package main; func main() { var ch chan int; _ = len(ch) }",
    zero_value_make_after_nil_compile => "package main; func main() { var values []int; values = make([]int, 2); _ = values }",
    zero_value_bool_negation_compile => "package main; func main() { var flag bool; _ = !flag }",
    nil_pointer_in_method_compile => "package main; type counter struct{}; func (c *counter) ok() bool { return c == nil }; func main() { var c *counter; _ = c.ok() }",
    zero_value_struct_with_array_compile => "package main; type holder struct { values [3]int }; func main() { var h holder; _ = h }",
    zero_value_struct_with_map_compile => "package main; type holder struct { values map[string]int }; func main() { var h holder; _ = h }",
    zero_value_struct_with_slice_compile => "package main; type holder struct { values []int }; func main() { var h holder; _ = h }",
    nil_interface_in_map_compile => "package main; func main() { values := map[string]interface{}{\"x\": nil}; _ = values }",
    nil_channel_select_compile => "package main; func main() { var ch chan int; select { case <-ch: default: } }",
}