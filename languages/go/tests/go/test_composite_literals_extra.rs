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

go_compile_test!(
    array_literal_with_indexed_elements_runtime,
    "package main; import \"fmt\"; func main() { values := [4]int{1: 7, 3: 9}; fmt.Println(values[1]); fmt.Println(values[3]); }"
);

go_run_test!(
    array_literal_inferred_length_runtime,
    "package main; import \"fmt\"; func main() { values := [...]int{2, 4, 6}; fmt.Println(len(values)); }",
    vec!["3"]
);

go_run_test!(
    array_literal_nested_arrays_runtime,
    "package main; import \"fmt\"; func main() { grid := [2][2]int{{1, 2}, {3, 4}}; fmt.Println(grid[1][0]); }",
    vec!["3"]
);

go_run_test!(
    array_literal_of_structs_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; func main() { pts := [2]point{{x: 1, y: 2}, {x: 3, y: 4}}; fmt.Println(pts[1].y); }",
    vec!["4"]
);

go_compile_test!(
    array_literal_of_pointers_compile,
    "package main; func main() { a, b := 1, 2; values := [2]*int{&a, &b}; _ = values }"
);

go_compile_test!(
    slice_literal_trailing_comma_compile,
    "package main; func main() { values := []int{1, 2, 3 }; _ = values }"
);

go_run_test!(
    slice_literal_nested_slices_runtime,
    "package main; import \"fmt\"; func main() { grid := [][]int{{1, 2}, {3, 4, 5}}; fmt.Println(grid[1][2]); }",
    vec!["5"]
);

go_run_test!(
    slice_literal_of_structs_runtime,
    "package main; import \"fmt\"; type user struct { name string }; func main() { users := []user{{name: \"a\"}, {name: \"b\"}}; fmt.Println(users[0].name); }",
    vec!["a"]
);

go_run_test!(
    slice_literal_with_named_fields_runtime,
    "package main; import \"fmt\"; type pair struct { left int; right int }; func main() { values := []pair{{left: 2, right: 3}, {left: 4, right: 5}}; fmt.Println(values[1].left + values[1].right); }",
    vec!["9"]
);

go_run_test!(
    map_literal_string_to_slice_runtime,
    "package main; import \"fmt\"; func main() { values := map[string][]int{\"odd\": {1, 3, 5}}; fmt.Println(values[\"odd\"][2]); }",
    vec!["5"]
);

go_run_test!(
    map_literal_string_to_struct_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; func main() { values := map[string]point{\"p\": {x: 8, y: 9}}; fmt.Println(values[\"p\"].x); }",
    vec!["8"]
);

go_run_test!(
    map_literal_nested_maps_runtime,
    "package main; import \"fmt\"; func main() { values := map[string]map[string]int{\"outer\": {\"inner\": 11}}; fmt.Println(values[\"outer\"][\"inner\"]); }",
    vec!["11"]
);

go_run_test!(
    map_literal_bool_keys_runtime,
    "package main; import \"fmt\"; func main() { flags := map[bool]string{true: \"on\", false: \"off\"}; fmt.Println(flags[true]); }",
    vec!["on"]
);

go_run_test!(
    struct_literal_keyed_order_independent_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; func main() { p := point{y: 4, x: 6}; fmt.Println(p.x - p.y); }",
    vec!["2"]
);

go_run_test!(
    struct_literal_unkeyed_ordered_runtime,
    "package main; import \"fmt\"; type pair struct { a int; b int }; func main() { p := pair{3, 7}; fmt.Println(p.b); }",
    vec!["7"]
);

go_run_test!(
    anonymous_struct_literal_runtime,
    "package main; import \"fmt\"; func main() { value := struct { label string; count int }{label: \"go\", count: 5}; fmt.Println(value.count); }",
    vec!["5"]
);

go_run_test!(
    pointer_to_struct_literal_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; func main() { p := &point{x: 5, y: 6}; fmt.Println(p.y); }",
    vec!["6"]
);

go_compile_test!(
    pointer_to_array_literal_compile,
    "package main; func main() { values := &[3]int{1, 2, 3}; _ = values }"
);

go_run_test!(
    slice_of_map_literals_runtime,
    "package main; import \"fmt\"; func main() { values := []map[string]int{{\"a\": 1}, {\"b\": 2}}; fmt.Println(values[1][\"b\"]); }",
    vec!["2"]
);

go_run_test!(
    map_of_slice_literals_runtime,
    "package main; import \"fmt\"; func main() { values := map[string][]string{\"langs\": {\"go\", \"rust\"}}; fmt.Println(values[\"langs\"][0]); }",
    vec!["go"]
);

go_run_test!(
    composite_literal_inside_return_runtime,
    "package main; import \"fmt\"; type pair struct { a int; b int }; func build() pair { return pair{a: 2, b: 9} }; func main() { value := build(); fmt.Println(value.b); }",
    vec!["9"]
);

go_run_test!(
    composite_literal_inside_function_arg_runtime,
    "package main; import \"fmt\"; type pair struct { a int; b int }; func sum(p pair) int { return p.a + p.b }; func main() { fmt.Println(sum(pair{a: 4, b: 5})); }",
    vec!["9"]
);

go_compile_test!(
    struct_literal_embedded_field_compile,
    "package main; type inner struct { value int }; type outer struct { inner }; func main() { _ = outer{inner: inner{value: 3}} }"
);

go_run_test!(
    nested_composite_literal_three_levels_runtime,
    "package main; import \"fmt\"; type cell struct { value int }; type row struct { cells []cell }; type table struct { rows []row }; func main() { t := table{rows: []row{{cells: []cell{{value: 8}}}}}; fmt.Println(t.rows[0].cells[0].value); }",
    vec!["8"]
);

go_run_test!(
    empty_slice_literal_len_runtime,
    "package main; import \"fmt\"; func main() { values := []int{}; fmt.Println(len(values)); }",
    vec!["0"]
);

go_run_test!(
    empty_map_literal_len_runtime,
    "package main; import \"fmt\"; func main() { values := map[string]int{}; fmt.Println(len(values)); }",
    vec!["0"]
);

go_run_test!(
    nil_slice_vs_empty_slice_runtime,
    "package main; import \"fmt\"; func main() { var a []int; b := []int{}; fmt.Println(len(a)); fmt.Println(len(b)); }",
    vec!["0", "0"]
);

go_run_test!(
    append_composite_struct_literal_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; func main() { values := []point{}; values = append(values, point{x: 2, y: 3}); fmt.Println(values[0].x + values[0].y); }",
    vec!["5"]
);

go_compile_test!(
    composite_literal_with_const_keys_compile,
    "package main; const home = \"home\"; func main() { values := map[string]int{home: 1}; _ = values }"
);

go_run_test!(
    keyed_array_literal_sparse_runtime,
    "package main; import \"fmt\"; func main() { values := [5]int{2: 7, 4: 9}; fmt.Println(values[2]); fmt.Println(values[4]); }",
    vec!["7", "9"]
);

go_run_test!(
    slice_literal_from_array_slice_runtime,
    "package main; import \"fmt\"; func main() { values := [4]int{1, 2, 3, 4}; part := values[1:3]; fmt.Println(part[0]); fmt.Println(part[1]); }",
    vec!["2", "3"]
);

go_run_test!(
    struct_literal_with_slice_field_runtime,
    "package main; import \"fmt\"; type bag struct { items []string }; func main() { b := bag{items: []string{\"a\", \"b\"}}; fmt.Println(b.items[1]); }",
    vec!["b"]
);

go_run_test!(
    struct_literal_with_map_field_runtime,
    "package main; import \"fmt\"; type config struct { values map[string]int }; func main() { c := config{values: map[string]int{\"size\": 4}}; fmt.Println(c.values[\"size\"]); }",
    vec!["4"]
);

go_run_test!(
    struct_literal_with_pointer_field_runtime,
    "package main; import \"fmt\"; type node struct { value int }; type wrapper struct { item *node }; func main() { n := node{value: 12}; w := wrapper{item: &n}; fmt.Println(w.item.value); }",
    vec!["12"]
);

go_compile_test!(
    slice_literal_with_interface_values_compile,
    "package main; func main() { values := []interface{}{1, \"two\", true}; _ = values }"
);

go_compile_test!(
    map_literal_with_interface_values_compile,
    "package main; func main() { values := map[string]interface{}{\"n\": 1, \"s\": \"two\"}; _ = values }"
);

go_run_test!(
    composite_literal_in_short_if_init_runtime,
    "package main; import \"fmt\"; type point struct { x int }; func main() { if p := point{x: 9}; p.x > 0 { fmt.Println(p.x) } }",
    vec!["9"]
);

go_run_test!(
    composite_literal_in_switch_init_runtime,
    "package main; import \"fmt\"; type pair struct { a int }; func main() { switch p := pair{a: 5}; p.a { case 5: fmt.Println(\"five\") } }",
    vec!["five"]
);

go_compile_test!(
    nested_anonymous_struct_literal_compile,
    "package main; func main() { value := struct { inner struct { n int } }{}; _ = value }"
);

go_run_test!(
    array_of_anonymous_structs_runtime,
    "package main; import \"fmt\"; func main() { values := [2]struct { n int }{{n: 2}, {n: 4}}; fmt.Println(values[1].n); }",
    vec!["4"]
);

go_run_test!(
    slice_of_anonymous_structs_runtime,
    "package main; import \"fmt\"; func main() { values := []struct { n int }{{n: 3}, {n: 6}}; fmt.Println(values[0].n + values[1].n); }",
    vec!["9"]
);

go_compile_test!(
    map_of_anonymous_structs_compile,
    "package main; func main() { values := map[string]struct { n int }{\"x\": {n: 1}}; _ = values }"
);

go_run_test!(
    composite_literal_using_named_type_alias_runtime,
    "package main; import \"fmt\"; type scores []int; func main() { values := scores{4, 5, 6}; fmt.Println(values[2]); }",
    vec!["6"]
);

go_run_test!(
    slice_literal_zero_value_elements_runtime,
    "package main; import \"fmt\"; func main() { values := make([]int, 3); fmt.Println(values[0]); fmt.Println(values[2]); }",
    vec!["0", "0"]
);

go_run_test!(
    map_literal_read_missing_after_literal_runtime,
    "package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; fmt.Println(values[\"missing\"]); }",
    vec!["0"]
);

go_run_test!(
    nested_keyed_struct_literal_runtime,
    "package main; import \"fmt\"; type address struct { city string }; type user struct { name string; addr address }; func main() { u := user{name: \"Ada\", addr: address{city: \"Rome\"}}; fmt.Println(u.addr.city); }",
    vec!["Rome"]
);

go_run_test!(
    composite_literal_returning_pointer_runtime,
    "package main; import \"fmt\"; type point struct { x int }; func build() *point { return &point{x: 13} }; func main() { fmt.Println(build().x); }",
    vec!["13"]
);

go_compile_test!(
    composite_literal_assigned_to_interface_compile,
    "package main; type any interface{}; func main() { var v any = struct { n int }{n: 7}; _ = v }"
);

go_run_test!(
    struct_literal_with_array_field_runtime,
    "package main; import \"fmt\"; type box struct { values [2]int }; func main() { b := box{values: [2]int{8, 9}}; fmt.Println(b.values[0]); }",
    vec!["8"]
);

go_run_test!(
    map_literal_integer_keys_runtime,
    "package main; import \"fmt\"; func main() { values := map[int]string{1: \"one\", 2: \"two\"}; fmt.Println(values[2]); }",
    vec!["two"]
);

go_run_test!(
    slice_of_pointer_literals_runtime,
    "package main; import \"fmt\"; type point struct { x int }; func main() { a := &point{x: 1}; b := &point{x: 2}; values := []*point{a, b}; fmt.Println(values[1].x); }",
    vec!["2"]
);

go_compile_test!(
    composite_literal_multiline_compile,
    "package main; type point struct { x int; y int }; func main() { _ = point{\n x: 1,\n y: 2,\n } }"
);
