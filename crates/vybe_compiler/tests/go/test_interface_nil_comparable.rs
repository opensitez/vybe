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
    // --- typed nil inside an interface is not interface-nil ---
    typed_nil_pointer_in_empty_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var p *int; var value interface{} = p; fmt.Println(value == nil) }", vec!["false"]),
    typed_nil_pointer_in_named_interface_not_nil =>
        ("package main; import \"fmt\"; type holder interface { size() int }; type box struct { n int }; func (b *box) size() int { return b.n }; func main() { var p *box; var value holder = p; fmt.Println(value == nil) }", vec!["false"]),
    typed_nil_slice_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var s []int; var value interface{} = s; fmt.Println(value == nil) }", vec!["false"]),
    typed_nil_map_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var m map[string]int; var value interface{} = m; fmt.Println(value == nil) }", vec!["false"]),
    typed_nil_channel_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var ch chan int; var value interface{} = ch; fmt.Println(value == nil) }", vec!["false"]),
    typed_nil_function_in_interface_not_nil =>
        ("package main; import \"fmt\"; func main() { var fn func(); var value interface{} = fn; fmt.Println(value == nil) }", vec!["false"]),
    untyped_nil_assigned_to_empty_interface_is_nil =>
        ("package main; import \"fmt\"; func main() { var value interface{} = nil; fmt.Println(value == nil) }", vec!["true"]),
    typed_nil_concrete_error_pointer_not_nil =>
        ("package main; import \"fmt\"; type myError struct { msg string }; func (e *myError) Error() string { return e.msg }; func main() { var p *myError; var err error = p; fmt.Println(err == nil) }", vec!["false"]),

    // --- interface equality (dynamic type + value) ---
    two_nil_empty_interfaces_are_equal =>
        ("package main; import \"fmt\"; func main() { var left interface{}; var right interface{}; fmt.Println(left == right) }", vec!["true"]),
    interface_int_equality_same_value =>
        ("package main; import \"fmt\"; func main() { var left interface{} = 7; var right interface{} = 7; fmt.Println(left == right) }", vec!["true"]),
    interface_int_inequality_different_values =>
        ("package main; import \"fmt\"; func main() { var left interface{} = 3; var right interface{} = 4; fmt.Println(left == right) }", vec!["false"]),
    interface_string_equality_same_value =>
        ("package main; import \"fmt\"; func main() { var left interface{} = \"go\"; var right interface{} = \"go\"; fmt.Println(left == right) }", vec!["true"]),
    interface_different_dynamic_types_not_equal =>
        ("package main; import \"fmt\"; func main() { var left interface{} = 1; var right interface{} = \"1\"; fmt.Println(left == right) }", vec!["false"]),
    interface_struct_equality_same_fields =>
        ("package main; import \"fmt\"; type point struct { x int; y int }; func main() { var left interface{} = point{x: 1, y: 2}; var right interface{} = point{x: 1, y: 2}; fmt.Println(left == right) }", vec!["true"]),
    interface_struct_inequality_different_fields =>
        ("package main; import \"fmt\"; type point struct { x int; y int }; func main() { var left interface{} = point{x: 1, y: 2}; var right interface{} = point{x: 2, y: 1}; fmt.Println(left == right) }", vec!["false"]),
    typed_nil_pointers_in_interfaces_equal_same_type =>
        ("package main; import \"fmt\"; func main() { var p *int; var left interface{} = p; var right interface{} = p; fmt.Println(left == right) }", vec!["true"]),
    typed_nil_pointers_in_interfaces_differ_by_type =>
        ("package main; import \"fmt\"; func main() { var pi *int; var ps *string; var left interface{} = pi; var right interface{} = ps; fmt.Println(left == right) }", vec!["false"]),
    named_interface_concrete_value_equality =>
        ("package main; import \"fmt\"; type counter interface { count() int }; type tally struct { n int }; func (t tally) count() int { return t.n }; func main() { var left counter = tally{n: 5}; var right counter = tally{n: 5}; fmt.Println(left == right) }", vec!["true"]),
    interface_reset_to_nil_after_value =>
        ("package main; import \"fmt\"; func main() { var value interface{} = 9; fmt.Println(value == nil); value = nil; fmt.Println(value == nil) }", vec!["false", "true"]),

    // --- nil map / slice / channel direct comparisons ---
    nil_slice_equals_nil_slice =>
        ("package main; import \"fmt\"; func main() { var left []int; var right []int; fmt.Println(left == nil); fmt.Println(left == right) }", vec!["true", "true"]),
    make_zero_length_slice_not_equal_nil =>
        ("package main; import \"fmt\"; func main() { empty := make([]int, 0); var nilSlice []int; fmt.Println(empty == nil); fmt.Println(nilSlice == nil) }", vec!["false", "true"]),
    nil_map_equals_nil_map =>
        ("package main; import \"fmt\"; func main() { var left map[string]int; var right map[string]int; fmt.Println(left == nil); fmt.Println(left == right) }", vec!["true", "true"]),
    make_empty_map_not_equal_nil =>
        ("package main; import \"fmt\"; func main() { empty := make(map[string]int); var nilMap map[string]int; fmt.Println(empty == nil); fmt.Println(nilMap == nil) }", vec!["false", "true"]),
    nil_channel_equals_nil_channel =>
        ("package main; import \"fmt\"; func main() { var left chan int; var right chan int; fmt.Println(left == nil); fmt.Println(left == right) }", vec!["true", "true"]),
    make_channel_not_equal_nil =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int); fmt.Println(ch == nil) }", vec!["false"]),
    interface_holding_nil_slice_vs_interface_nil =>
        ("package main; import \"fmt\"; func main() { var s []int; var boxed interface{} = s; var empty interface{}; fmt.Println(boxed == empty) }", vec!["false"]),

    // --- comparable constraint usage with nil and equality ---
    generic_comparable_int_equality =>
        ("package main; import \"fmt\"; func equal[T comparable](left T, right T) bool { return left == right }; func main() { fmt.Println(equal(3, 3)); fmt.Println(equal(3, 4)) }", vec!["true", "false"]),
    generic_comparable_nil_pointer_param =>
        ("package main; import \"fmt\"; func isNil[T comparable](value T) bool { var zero T; return value == zero }; func main() { var p *int; fmt.Println(isNil(p)) }", vec!["true"]),
    generic_comparable_string_nil_map_key_lookup =>
        ("package main; import \"fmt\"; func lookup[K comparable, V any](m map[K]V, key K) bool { _, ok := m[key]; return ok }; func main() { var m map[string]int; fmt.Println(lookup(m, \"missing\")) }", vec!["false"]),
    generic_comparable_bool_zero_value =>
        ("package main; import \"fmt\"; func isZero[T comparable](value T) bool { var zero T; return value == zero }; func main() { var flag bool; fmt.Println(isZero(flag)) }", vec!["true"]),
    generic_comparable_array_equality =>
        ("package main; import \"fmt\"; func equalArray(left [2]int, right [2]int) bool { return left == right }; func main() { fmt.Println(equalArray([2]int{1, 2}, [2]int{1, 2})); fmt.Println(equalArray([2]int{1, 2}, [2]int{2, 1})) }", vec!["true", "false"]),
}

compile_cases! {
    interface_equality_compile =>
        "package main; func main() { var left interface{} = 1; var right interface{} = 1; _ = (left == right) }",
    typed_nil_interface_inequality_to_nil_compile =>
        "package main; func main() { var p *int; var value interface{} = p; _ = (value != nil) }",
    comparable_generic_map_key_compile =>
        "package main; func keys[K comparable, V any](m map[K]V) []K { result := make([]K, 0); for k := range m { result = append(result, k) }; return result }; func main() { _ = keys(map[int]string{}) }",
    nil_slice_compare_nil_compile =>
        "package main; func main() { var values []int; _ = (values == nil) }",
    nil_map_compare_nil_compile =>
        "package main; func main() { var values map[string]int; _ = (values == nil) }",
    nil_channel_compare_nil_compile =>
        "package main; func main() { var ch chan int; _ = (ch == nil) }",
    named_interface_typed_nil_parameter_compile =>
        "package main; type reader interface { read() int }; type book struct{}; func (b *book) read() int { return 1 }; func accept(value reader) bool { return value == nil }; func main() { var p *book; _ = accept(p) }",
    interface_compare_after_type_assertion_compile =>
        "package main; func main() { var value interface{} = 2; number, ok := value.(int); if ok { _ = (number == 2) } }",
}
