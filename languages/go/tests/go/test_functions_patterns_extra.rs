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
    immediately_invoked_function_runtime => ("package main; import \"fmt\"; func main() { value := func(n int) int { return n * 2 }(6); fmt.Println(value); }", vec!["12"]),
    function_returning_function_runtime => ("package main; import \"fmt\"; func maker(step int) func(int) int { return func(v int) int { return v + step } }; func main() { next := maker(3); fmt.Println(next(4)); }", vec!["7"]),
    closure_mutates_captured_local_runtime => ("package main; import \"fmt\"; func main() { total := 0; add := func(v int) { total += v }; add(2); add(5); fmt.Println(total); }", vec!["7"]),
    callback_applies_operation_runtime => ("package main; import \"fmt\"; func apply(v int, fn func(int) int) int { return fn(v) }; func main() { result := apply(4, func(v int) int { return v * v }); fmt.Println(result); }", vec!["16"]),
    variadic_forwarding_runtime => ("package main; import \"fmt\"; func sum(values ...int) int { total := 0; for _, v := range values { total += v }; return total }; func wrap(values ...int) int { return sum(values...) }; func main() { fmt.Println(wrap(1, 2, 3)); }", vec!["6"]),
    named_return_bare_return_runtime => ("package main; import \"fmt\"; func twice(v int) (result int) { result = v * 2; return }; func main() { fmt.Println(twice(5)); }", vec!["10"]),
    named_return_explicit_assignment_runtime => ("package main; import \"fmt\"; func classify(v int) (label string) { if v > 0 { label = \"pos\"; return }; label = \"zero\"; return }; func main() { fmt.Println(classify(3)); }", vec!["pos"]),
    tuple_return_used_in_if_runtime => ("package main; import \"fmt\"; func dims() (int, int) { return 2, 4 }; func main() { w, h := dims(); if w < h { fmt.Println(h - w) } }", vec!["2"]),
    function_value_reassignment_runtime => ("package main; import \"fmt\"; func add(a int, b int) int { return a + b }; func mul(a int, b int) int { return a * b }; func main() { op := add; fmt.Println(op(2, 3)); op = mul; fmt.Println(op(2, 3)); }", vec!["5", "6"]),
    anonymous_func_as_argument_runtime => ("package main; import \"fmt\"; func consume(fn func() string) string { return fn() }; func main() { fmt.Println(consume(func() string { return \"ok\" })); }", vec!["ok"]),
    return_function_from_switch_runtime => ("package main; import \"fmt\"; func choose(flag bool) func(int) int { switch flag { case true: return func(v int) int { return v + 1 }; default: return func(v int) int { return v - 1 } } }; func main() { fmt.Println(choose(true)(8)); }", vec!["9"]),
    closure_reads_outer_after_mutation_runtime => ("package main; import \"fmt\"; func main() { prefix := \"go\"; fn := func() string { return prefix }; prefix = \"vybe\"; fmt.Println(fn()); }", vec!["vybe"]),
    variadic_sum_with_slice_expansion_runtime => ("package main; import \"fmt\"; func sum(values ...int) int { total := 0; for _, v := range values { total += v }; return total }; func main() { nums := []int{4, 5, 6}; fmt.Println(sum(nums...)); }", vec!["15"]),
    function_literal_in_map_runtime => ("package main; import \"fmt\"; func main() { ops := map[string]func(int) int{\"inc\": func(v int) int { return v + 1 }}; fmt.Println(ops[\"inc\"](9)); }", vec!["10"]),
    function_literal_in_struct_field_runtime => ("package main; import \"fmt\"; type holder struct { fn func(int) int }; func main() { h := holder{fn: func(v int) int { return v * 3 }}; fmt.Println(h.fn(4)); }", vec!["12"]),
    function_receives_function_return_runtime => ("package main; import \"fmt\"; func builder() func(int) int { return func(v int) int { return v + 2 } }; func apply(v int, fn func(int) int) int { return fn(v) }; func main() { fmt.Println(apply(5, builder())); }", vec!["7"]),
    higher_order_pipeline_runtime => ("package main; import \"fmt\"; func pipe(v int, a func(int) int, b func(int) int) int { return b(a(v)) }; func main() { fmt.Println(pipe(3, func(v int) int { return v + 1 }, func(v int) int { return v * 2 })); }", vec!["8"]),
    tuple_return_ignored_second_runtime => ("package main; import \"fmt\"; func pair() (int, string) { return 8, \"unused\" }; func main() { value, _ := pair(); fmt.Println(value); }", vec!["8"]),
    named_return_with_early_if_runtime => ("package main; import \"fmt\"; func abs(v int) (result int) { if v < 0 { result = -v; return }; result = v; return }; func main() { fmt.Println(abs(-4)); }", vec!["4"]),
    function_value_in_slice_runtime => ("package main; import \"fmt\"; func main() { fns := []func(int) int{func(v int) int { return v + 1 }, func(v int) int { return v + 2 }}; fmt.Println(fns[1](5)); }", vec!["7"]),
    function_literal_returns_struct_runtime => ("package main; import \"fmt\"; type pair struct { a int; b int }; func main() { build := func() pair { return pair{a: 3, b: 4} }; value := build(); fmt.Println(value.a + value.b); }", vec!["7"]),
    anonymous_function_immediate_args_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(func(a int, b int) int { return a - b }(9, 4)); }", vec!["5"]),
    function_composition_runtime => ("package main; import \"fmt\"; func compose(a func(int) int, b func(int) int) func(int) int { return func(v int) int { return b(a(v)) } }; func main() { fn := compose(func(v int) int { return v + 1 }, func(v int) int { return v * 2 }); fmt.Println(fn(5)); }", vec!["12"]),
    function_variable_default_nil_check_runtime => ("package main; import \"fmt\"; func main() { var fn func(int) int; fmt.Println(fn == nil); }", vec!["true"]),
    named_func_type_runtime => ("package main; import \"fmt\"; type op func(int, int) int; func main() { var add op = func(a int, b int) int { return a + b }; fmt.Println(add(2, 8)); }", vec!["10"]) }

compile_cases! {
    named_return_with_defer_compile => "package main; func run() (result int) { defer func() { result++ }(); result = 1; return }; func main() { _ = run }",
    function_type_parameter_compile => "package main; type transformer func(int) int; func apply(v int, fn transformer) int { return fn(v) }; func main() { _ = apply }",
    function_returning_named_type_compile => "package main; type score int; func build() score { return 5 }; func main() { _ = build() }",
    variadic_parameter_without_use_compile => "package main; func log(values ...int) {}; func main() { log() }",
    nested_function_literals_compile => "package main; func main() { outer := func() func() int { return func() int { return 1 } }; _ = outer }",
    function_value_nil_compare_compile => "package main; func main() { var fn func(); _ = (fn == nil) }",
    closure_over_loop_var_compile => "package main; func main() { fns := []func() int{}; for i := 0; i < 2; i++ { fns = append(fns, func() int { return i }) }; _ = fns }",
    function_type_alias_map_value_compile => "package main; type reducer func(int, int) int; func main() { ops := map[string]reducer{}; _ = ops }",
    function_type_alias_struct_field_compile => "package main; type reducer func(int) int; type holder struct { fn reducer }; func main() { _ = holder{} }",
    variadic_of_interface_compile => "package main; func pack(values ...interface{}) []interface{} { return values }; func main() { _ = pack(1, \"two\", true) }",
    multiple_named_returns_compile => "package main; func split(v int) (left int, right int) { left = v; right = v + 1; return }; func main() { _, _ = split(1) }",
    function_literal_with_named_result_compile => "package main; func main() { fn := func(v int) (result int) { result = v + 1; return }; _ = fn }",
    deferred_function_literal_compile => "package main; func main() { defer func() { _ = 1 }() }",
    function_param_shadow_compile => "package main; func use(v int) int { { v := v + 1; _ = v }; return v }; func main() { _ = use(1) }",
    function_value_in_array_compile => "package main; func main() { values := [1]func(){func() {}}; _ = values }",
    function_literal_assigned_later_compile => "package main; func main() { var fn func(int) int; fn = func(v int) int { return v }; _ = fn }",
    named_result_with_short_decl_compile => "package main; func run() (result int) { if value := 3; value > 0 { result = value }; return }; func main() { _ = run }",
    function_returning_pointer_compile => "package main; type point struct { x int }; func build() *point { return &point{x: 1} }; func main() { _ = build() }",
    anonymous_func_in_if_compile => "package main; func main() { if func() bool { return true }() { _ = 1 } }",
    anonymous_func_in_switch_compile => "package main; func main() { switch func() int { return 2 }() { case 2: _ = 2 } }",
    higher_order_returning_variadic_compile => "package main; func build() func(...int) int { return func(values ...int) int { return len(values) } }; func main() { _ = build() }",
    variadic_pass_zero_arguments_compile => "package main; func sum(values ...int) int { return len(values) }; func main() { _ = sum() }",
    function_array_return_compile => "package main; func build() [2]int { return [2]int{1, 2} }; func main() { _ = build() }",
    function_literal_with_blank_identifier_param_compile => "package main; func main() { fn := func(_ int, v int) int { return v }; _ = fn }",
    function_parameter_named_result_same_type_compile => "package main; func transform(v int) (int, int) { return v, v + 1 }; func main() { _, _ = transform(1) }" }
