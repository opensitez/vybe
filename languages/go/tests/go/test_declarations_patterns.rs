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

go_run_test!(
    package_level_var_runtime,
    "package main; import \"fmt\"; var greeting = \"hello\"; func main() { fmt.Println(greeting); }",
    vec!["hello"]
);

go_run_test!(
    package_level_const_runtime,
    "package main; import \"fmt\"; const answer = 42; func main() { fmt.Println(answer); }",
    vec!["42"]
);

go_compile_test!(
    grouped_var_block_compile,
    "package main; func main() { var ( a int = 1; b string = \"two\"; c bool ); _, _, _ = a, b, c }"
);

go_compile_test!(
    grouped_const_block_compile,
    "package main; const ( low = 1; high = 9 ); func main() { _ = low; _ = high }"
);

go_compile_test!(
    grouped_type_block_compile,
    "package main; type ( Score int; Label string ); func main() { var s Score = 3; var l Label = \"ok\"; _, _ = s, l }"
);

go_run_test!(
    package_scope_type_alias_runtime,
    "package main; import \"fmt\"; type Counter int; var total Counter = 7; func main() { fmt.Println(total); }",
    vec!["7"]
);

go_run_test!(
    init_function_sets_global_runtime,
    "package main; import \"fmt\"; var total int; func init() { total = 9 }; func main() { fmt.Println(total); }",
    vec!["9"]
);

go_compile_test!(
    multiple_init_functions_compile,
    "package main; var a int; func init() { a = 1 }; func init() { a = a + 1 }; func main() { _ = a }"
);

go_run_test!(
    blank_identifier_discards_return_runtime,
    "package main; import \"fmt\"; func pair() (int, int) { return 3, 4 }; func main() { x, _ := pair(); fmt.Println(x); }",
    vec!["3"]
);

go_run_test!(
    short_declaration_reuses_existing_runtime,
    "package main; import \"fmt\"; func main() { x := 1; { x, y := x+1, 8; fmt.Println(x); fmt.Println(y); }; fmt.Println(x); }",
    vec!["2", "8", "1"]
);

go_run_test!(
    short_declaration_in_if_init_runtime,
    "package main; import \"fmt\"; func main() { if n := 3 * 3; n > 5 { fmt.Println(n) } }",
    vec!["9"]
);

go_run_test!(
    short_declaration_in_switch_init_runtime,
    "package main; import \"fmt\"; func main() { switch n := 2 + 3; n { case 5: fmt.Println(\"five\") } }",
    vec!["five"]
);

go_run_test!(
    multi_assignment_swap_runtime,
    "package main; import \"fmt\"; func main() { a, b := 1, 2; a, b = b, a; fmt.Println(a); fmt.Println(b); }",
    vec!["2", "1"]
);

go_run_test!(
    multi_assignment_with_blank_identifier_runtime,
    "package main; import \"fmt\"; func triple() (int, int, int) { return 4, 5, 6 }; func main() { a, _, c := triple(); fmt.Println(a); fmt.Println(c); }",
    vec!["4", "6"]
);

go_compile_test!(
    var_block_mixed_initializers_compile,
    "package main; func main() { var ( a = 1; b int; c = a + 2 ); _, _, _ = a, b, c }"
);

go_compile_test!(
    const_block_mixed_iota_compile,
    "package main; const ( first = iota; second; third ); func main() { _, _, _ = first, second, third }"
);

go_run_test!(
    nested_short_declaration_in_inner_scope_runtime,
    "package main; import \"fmt\"; func main() { value := 10; { value := value + 5; fmt.Println(value) }; fmt.Println(value) }",
    vec!["15", "10"]
);

go_compile_test!(
    package_level_function_var_compile,
    "package main; func add(a int, b int) int { return a + b }; var op func(int, int) int = add; func main() { _ = op }"
);

go_run_test!(
    zero_value_package_var_runtime,
    "package main; import \"fmt\"; var total int; func main() { fmt.Println(total); }",
    vec!["0"]
);

go_compile_test!(
    grouped_imports_compile,
    "package main; import ( \"fmt\"; \"strings\" ); func main() { _, _ = fmt.Sprintf(\"%s\", \"x\"), strings.HasPrefix(\"go\", \"g\") }"
);

go_compile_test!(
    parenthesized_single_import_compile,
    "package main; import ( \"fmt\" ); func main() { _ = fmt.Sprintf(\"%d\", 3) }"
);

go_compile_test!(
    alias_import_compile,
    "package main; import f \"fmt\"; func main() { _ = f.Sprintf(\"%s\", \"alias\") }"
);

go_compile_test!(
    blank_import_compile,
    "package main; import _ \"fmt\"; func main() {}"
);

go_compile_test!(
    exported_identifier_compile,
    "package main; type Person struct { Name string }; func main() { _ = Person{Name: \"Ada\"} }"
);

go_run_test!(
    package_var_depends_on_const_runtime,
    "package main; import \"fmt\"; const base = 5; var total = base * 2; func main() { fmt.Println(total); }",
    vec!["10"]
);

go_compile_test!(
    local_type_inside_function_compile,
    "package main; func main() { type local struct { value int }; v := local{value: 3}; _ = v }"
);

go_run_test!(
    multiple_short_decl_with_function_call_runtime,
    "package main; import \"fmt\"; func pair() (int, int) { return 8, 9 }; func main() { a, b := pair(); fmt.Println(a + b); }",
    vec!["17"]
);

go_run_test!(
    declaration_then_tuple_reassignment_runtime,
    "package main; import \"fmt\"; func main() { a, b := 1, 2; a, b = a+b, a*b; fmt.Println(a); fmt.Println(b); }",
    vec!["3", "2"]
);

go_run_test!(
    tuple_assignment_from_multi_return_runtime,
    "package main; import \"fmt\"; func dims() (int, int) { return 4, 6 }; func main() { w, h := dims(); fmt.Println(w * h); }",
    vec!["24"]
);

go_compile_test!(
    short_declaration_err_style_compile,
    "package main; type simpleErr struct{}; func (simpleErr) Error() string { return \"err\" }; func pair() (int, error) { return 1, simpleErr{} }; func main() { n, err := pair(); _, _ = n, err }"
);

go_run_test!(
    shadow_package_var_in_function_runtime,
    "package main; import \"fmt\"; var name = \"pkg\"; func main() { name := \"local\"; fmt.Println(name); fmt.Println(name != \"pkg\") }",
    vec!["local", "true"]
);

go_run_test!(
    shadow_parameter_with_short_decl_inner_scope_runtime,
    "package main; import \"fmt\"; func bump(x int) int { if true { x := x + 1; return x }; return x }; func main() { fmt.Println(bump(2)); }",
    vec!["3"]
);

go_compile_test!(
    declaration_order_with_types_compile,
    "package main; type meter int; var distance meter = 4; func main() { _ = distance }"
);

go_compile_test!(
    var_block_with_inferred_types_compile,
    "package main; func main() { var ( a = 1; b = \"two\"; c = true ); _, _, _ = a, b, c }"
);

go_compile_test!(
    const_block_with_explicit_types_compile,
    "package main; const ( a int = 1; b string = \"two\" ); func main() { _, _ = a, b }"
);

go_run_test!(
    package_level_array_var_runtime,
    "package main; import \"fmt\"; var values = [3]int{1, 2, 3}; func main() { fmt.Println(values[2]); }",
    vec!["3"]
);

go_run_test!(
    package_level_struct_var_runtime,
    "package main; import \"fmt\"; type point struct { x int; y int }; var origin = point{x: 4, y: 5}; func main() { fmt.Println(origin.x + origin.y); }",
    vec!["9"]
);

go_compile_test!(
    package_level_map_var_compile,
    "package main; var lookup = map[string]int{\"go\": 1}; func main() { _ = lookup }"
);

go_compile_test!(
    package_level_slice_var_compile,
    "package main; var ids = []int{1, 2, 3}; func main() { _ = ids }"
);

go_run_test!(
    init_reads_package_var_runtime,
    "package main; import \"fmt\"; var base = 6; var total int; func init() { total = base + 1 }; func main() { fmt.Println(total); }",
    vec!["7"]
);

go_run_test!(
    package_const_expression_runtime,
    "package main; import \"fmt\"; const minutesPerHour = 60; const totalMinutes = minutesPerHour * 2; func main() { fmt.Println(totalMinutes); }",
    vec!["120"]
);

go_run_test!(
    local_const_group_runtime,
    "package main; import \"fmt\"; func main() { const ( a = 2; b = 3 ); fmt.Println(a * b); }",
    vec!["6"]
);

go_run_test!(
    local_var_group_runtime,
    "package main; import \"fmt\"; func main() { var ( a = 2; b = 4 ); fmt.Println(a + b); }",
    vec!["6"]
);

go_compile_test!(
    nested_var_block_compile,
    "package main; func main() { var a int = 1; { var b int = a + 1; _ = b }; _ = a }"
);

go_run_test!(
    explicit_zero_values_in_var_block_runtime,
    "package main; import \"fmt\"; func main() { var ( a int; b bool; c string ); fmt.Println(a); fmt.Println(b); fmt.Println(c == \"\"); }",
    vec!["0", "false", "true"]
);

go_run_test!(
    package_level_bool_var_runtime,
    "package main; import \"fmt\"; var enabled = true; func main() { fmt.Println(enabled); }",
    vec!["true"]
);

go_run_test!(
    package_level_string_var_runtime,
    "package main; import \"fmt\"; var title = \"vybe\"; func main() { fmt.Println(title); }",
    vec!["vybe"]
);

go_compile_test!(
    package_level_func_literal_compile,
    "package main; var op = func(a int, b int) int { return a + b }; func main() { _ = op }"
);

go_run_test!(
    package_level_named_function_value_runtime,
    "package main; import \"fmt\"; func add(a int, b int) int { return a + b }; var op = add; func main() { fmt.Println(op(2, 5)); }",
    vec!["7"]
);

go_compile_test!(
    grouped_declarations_with_comments_compile,
    "package main; func main() { var ( a = 1; // one\n b = 2; // two\n ); _ = a + b }"
);

go_run_test!(
    package_level_iota_runtime,
    "package main; import \"fmt\"; const ( red = iota; green; blue ); func main() { fmt.Println(red); fmt.Println(green); fmt.Println(blue); }",
    vec!["0", "1", "2"]
);

go_compile_test!(
    type_declaration_inside_const_scope_compile,
    "package main; type score int; const top score = 9; func main() { _ = top }"
);
