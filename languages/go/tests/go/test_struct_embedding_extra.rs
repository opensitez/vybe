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
    embedded_field_explicit_access_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { value := outer{inner: inner{count: 7}}; fmt.Println(value.inner.count); }", vec!["7"]),
    embedded_field_promotion_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { value := outer{inner: inner{count: 7}}; fmt.Println(value.count); }", vec!["7"]),
    embedded_method_promotion_runtime => ("package main; import \"fmt\"; type inner struct{}; func (inner) label() string { return \"ok\" }; type outer struct { inner }; func main() { value := outer{}; fmt.Println(value.label()); }", vec!["ok"]),
    embedded_field_shadow_runtime => ("package main; import \"fmt\"; type inner struct { name string }; type outer struct { inner; name string }; func main() { value := outer{inner: inner{name: \"inner\"}, name: \"outer\"}; fmt.Println(value.name); fmt.Println(value.inner.name); }", vec!["outer", "inner"]),
    zero_value_embedded_struct_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type outer struct { inner }; func main() { var value outer; fmt.Println(value.count); }", vec!["0"]),
    nested_struct_literal_field_runtime => ("package main; import \"fmt\"; type point struct { x int; y int }; type box struct { p point }; func main() { value := box{p: point{x: 3, y: 4}}; fmt.Println(value.p.x + value.p.y); }", vec!["7"]),
    struct_assignment_copies_value_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func main() { left := counter{n: 3}; right := left; right.n = 8; fmt.Println(left.n); fmt.Println(right.n); }", vec!["3", "8"]),
    struct_pass_to_function_runtime => ("package main; import \"fmt\"; type point struct { x int; y int }; func total(p point) int { return p.x + p.y }; func main() { fmt.Println(total(point{x: 2, y: 5})); }", vec!["7"]),
    struct_return_from_function_runtime => ("package main; import \"fmt\"; type point struct { x int; y int }; func build() point { return point{x: 4, y: 6} }; func main() { value := build(); fmt.Println(value.x); fmt.Println(value.y); }", vec!["4", "6"]),
    struct_array_field_runtime => ("package main; import \"fmt\"; type bag struct { values [3]int }; func main() { value := bag{values: [3]int{1, 2, 3}}; fmt.Println(value.values[2]); }", vec!["3"]),
    struct_slice_field_len_runtime => ("package main; import \"fmt\"; type bag struct { values []int }; func main() { value := bag{values: []int{2, 4, 6}}; fmt.Println(len(value.values)); }", vec!["3"]),
    struct_map_field_lookup_runtime => ("package main; import \"fmt\"; type bag struct { values map[string]int }; func main() { value := bag{values: map[string]int{\"x\": 9}}; fmt.Println(value.values[\"x\"]); }", vec!["9"]),
    anonymous_struct_literal_runtime => ("package main; import \"fmt\"; func main() { value := struct { left int; right int }{left: 2, right: 8}; fmt.Println(value.left + value.right); }", vec!["10"]),
    anonymous_struct_in_slice_runtime => ("package main; import \"fmt\"; func main() { values := []struct { name string }{{name: \"vybe\"}}; fmt.Println(values[0].name); }", vec!["vybe"]),
    struct_pointer_field_nil_runtime => ("package main; import \"fmt\"; type node struct { next *node }; func main() { value := node{}; fmt.Println(value.next == nil); }", vec!["true"]),
    struct_function_field_runtime => ("package main; import \"fmt\"; type holder struct { fn func(int) int }; func main() { value := holder{fn: func(v int) int { return v + 3 }}; fmt.Println(value.fn(4)); }", vec!["7"]),
    struct_in_map_value_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { values := map[string]point{\"a\": {x: 11}}; fmt.Println(values[\"a\"].x); }", vec!["11"]),
    struct_in_array_value_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { values := [2]point{{x: 1}, {x: 5}}; fmt.Println(values[1].x); }", vec!["5"]),
    struct_bool_field_branch_runtime => ("package main; import \"fmt\"; type state struct { ok bool }; func main() { value := state{ok: true}; if value.ok { fmt.Println(1) } else { fmt.Println(0) } }", vec!["1"]),
    struct_string_concat_runtime => ("package main; import \"fmt\"; type label struct { prefix string; suffix string }; func main() { value := label{prefix: \"vy\", suffix: \"be\"}; fmt.Println(value.prefix + value.suffix); }", vec!["vybe"]),
    struct_partial_literal_defaults_runtime => ("package main; import \"fmt\"; type item struct { count int; name string }; func main() { value := item{name: \"go\"}; fmt.Println(value.count); fmt.Println(value.name); }", vec!["0", "go"]),
    embedded_nested_access_runtime => ("package main; import \"fmt\"; type inner struct { count int }; type middle struct { inner }; type outer struct { middle }; func main() { value := outer{middle: middle{inner: inner{count: 12}}}; fmt.Println(value.count); }", vec!["12"]),
    struct_swap_values_runtime => ("package main; import \"fmt\"; type point struct { x int }; func main() { left := point{x: 1}; right := point{x: 9}; left, right = right, left; fmt.Println(left.x); fmt.Println(right.x); }", vec!["9", "1"]),
    struct_copy_before_mutation_runtime => ("package main; import \"fmt\"; type item struct { n int }; func main() { original := item{n: 4}; copy := original; original.n = 10; fmt.Println(copy.n); fmt.Println(original.n); }", vec!["4", "10"]),
    nested_embedded_field_runtime => ("package main; import \"fmt\"; type flags struct { enabled bool }; type config struct { flags }; type app struct { config }; func main() { value := app{config: config{flags: flags{enabled: true}}}; fmt.Println(value.enabled); }", vec!["true"]),
    empty_struct_slice_len_runtime => ("package main; import \"fmt\"; type token struct{}; func main() { values := []token{{}, {}}; fmt.Println(len(values)); }", vec!["2"]),
}

compile_cases! {
    empty_struct_value_compile => "package main; type marker struct{}; func main() { _ = marker{} }",
    recursive_struct_pointer_compile => "package main; type node struct { next *node }; func main() { var n node; _ = n }",
    embedded_pointer_field_compile => "package main; type inner struct { count int }; type outer struct { *inner }; func main() { _ = outer{} }",
    anonymous_struct_type_compile => "package main; func main() { var value struct { count int; label string }; _ = value }",
    struct_array_literal_compile => "package main; type point struct { x int }; func main() { _ = [2]point{{x: 1}, {x: 2}} }",
    struct_slice_literal_compile => "package main; type point struct { x int }; func main() { _ = []point{{x: 1}, {x: 2}} }",
    struct_map_literal_compile => "package main; type point struct { x int }; func main() { _ = map[string]point{\"a\": {x: 1}} }",
    embedded_method_compile => "package main; type inner struct{}; func (inner) label() string { return \"ok\" }; type outer struct { inner }; func main() { var value outer; _ = value.label() }",
    struct_with_function_field_compile => "package main; type holder struct { fn func(int) int }; func main() { _ = holder{} }",
    struct_with_interface_field_compile => "package main; type holder struct { value interface{} }; func main() { _ = holder{} }",
    nested_struct_field_compile => "package main; type inner struct { count int }; type outer struct { value inner }; func main() { _ = outer{} }",
    struct_return_compile => "package main; type point struct { x int }; func build() point { return point{x: 1} }; func main() { _ = build() }",
    struct_parameter_compile => "package main; type point struct { x int }; func use(value point) int { return value.x }; func main() { _ = use }",
    anonymous_struct_slice_compile => "package main; func main() { values := []struct { n int }{{n: 1}}; _ = values }",
    anonymous_struct_map_compile => "package main; func main() { values := map[string]struct { n int }{\"a\": {n: 1}}; _ = values }",
    promoted_field_selector_compile => "package main; type inner struct { count int }; type outer struct { inner }; func main() { var value outer; _ = value.count }",
    nested_embedded_compile => "package main; type inner struct { count int }; type middle struct { inner }; type outer struct { middle }; func main() { var value outer; _ = value.count }",
    struct_literal_in_return_compile => "package main; type point struct { x int }; func build() point { return point{x: 2} }; func main() { _ = build }",
    struct_pointer_field_compile => "package main; type node struct { next *node }; func main() { _ = node{next: nil} }",
    struct_with_array_field_compile => "package main; type bag struct { values [2]int }; func main() { _ = bag{} }",
    struct_with_slice_field_compile => "package main; type bag struct { values []int }; func main() { _ = bag{} }",
    struct_with_map_field_compile => "package main; type bag struct { values map[string]int }; func main() { _ = bag{} }",
    struct_with_embedded_empty_compile => "package main; type marker struct{}; type holder struct { marker }; func main() { _ = holder{} }",
    struct_composite_assignment_compile => "package main; type point struct { x int }; func main() { var left point; left = point{x: 1}; _ = left }",
    struct_nested_literal_compile => "package main; type inner struct { count int }; type outer struct { value inner }; func main() { _ = outer{value: inner{count: 3}} }",
}
