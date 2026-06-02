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
    defer_named_return_add_two_runtime => ("package main; import \"fmt\"; func build() (result int) { defer func() { result += 2 }(); return 3 }; func main() { fmt.Println(build()); }", vec!["5"]),
    defer_mutate_pointer_param_runtime => ("package main; import \"fmt\"; func main() { value := 1; func() { defer func(ptr *int) { *ptr = 5 }(&value) }(); fmt.Println(value); }", vec!["5"]),
    defer_local_cleanup_runtime => ("package main; import \"fmt\"; func main() { done := 0; func() { defer func() { done = 3 }() }(); fmt.Println(done); }", vec!["3"]),
    recover_nil_without_panic_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(recover() == nil); }", vec!["true"]),
    nested_defer_lifo_with_closures_runtime => ("package main; import \"fmt\"; func main() { defer func() { fmt.Println(\"first\") }(); defer func() { fmt.Println(\"second\") }(); }", vec!["second", "first"]),
    panic_recover_allows_continue_runtime => ("package main; import \"fmt\"; func safe() { defer func() { recover() }(); panic(\"x\") }; func main() { safe(); fmt.Println(1); }", vec!["1"]),
    multiple_deferred_mutations_named_return_runtime => ("package main; import \"fmt\"; func build() (result int) { defer func() { result += 3 }(); defer func() { result *= 2 }(); result = 4; return }; func main() { fmt.Println(build()); }", vec!["11"]),
    recover_type_preserved_string_runtime => ("package main; import \"fmt\"; func run() { defer func() { value := recover(); fmt.Println(value == \"err\") }(); panic(\"err\") }; func main() { run() }", vec!["true"]),
    defer_order_with_named_functions_runtime => ("package main; import \"fmt\"; func one() { fmt.Println(1) }; func two() { fmt.Println(2) }; func main() { defer one(); defer two(); }", vec!["2", "1"]),
    defer_in_branch_runtime => ("package main; import \"fmt\"; func main() { if true { defer fmt.Println(2) }; fmt.Println(1); }", vec!["1", "2"]),
    defer_before_multiple_returns_runtime => ("package main; import \"fmt\"; func build(flag bool) int { defer fmt.Println(\"done\"); if flag { return 1 }; return 2 }; func main() { fmt.Println(build(false)); }", vec!["done", "2"]),
    defer_modify_slice_runtime => ("package main; import \"fmt\"; func main() { values := []int{1}; func() { defer func() { values[0] = 9 }() }(); fmt.Println(values[0]); }", vec!["9"]),
    defer_modify_map_runtime => ("package main; import \"fmt\"; func main() { values := map[string]int{\"a\": 1}; func() { defer func() { values[\"a\"] = 7 }() }(); fmt.Println(values[\"a\"]); }", vec!["7"]),
    defer_print_after_return_value_runtime => ("package main; import \"fmt\"; func build() int { defer fmt.Println(\"later\"); return 4 }; func main() { fmt.Println(build()); }", vec!["later", "4"]),
    defer_in_for_with_closure_capture_runtime => ("package main; import \"fmt\"; func main() { for i := 0; i < 2; i++ { value := i; defer func() { fmt.Println(value) }() } }", vec!["1", "0"]),
    panic_after_defer_runtime => ("package main; import \"fmt\"; func run() { defer fmt.Println(\"cleanup\"); defer func() { recover() }(); panic(\"stop\") }; func main() { run() }", vec!["cleanup"]),
    recover_result_reused_runtime => ("package main; import \"fmt\"; func run() { defer func() { value := recover(); fmt.Println(value); fmt.Println(value != nil) }(); panic(3) }; func main() { run() }", vec!["3", "true"]),
    defer_updates_outer_variable_runtime => ("package main; import \"fmt\"; func main() { total := 1; func() { defer func() { total = 8 }() }(); fmt.Println(total); }", vec!["8"]),
    defer_method_on_pointer_receiver_runtime => ("package main; import \"fmt\"; type counter struct { n int }; func (c *counter) show() { fmt.Println(c.n) }; func main() { value := &counter{n: 6}; defer value.show(); value.n = 9; }", vec!["9"]),
    recover_in_nested_function_runtime => ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover() != nil) }(); panic(\"boom\") }; func main() { run() }", vec!["true"]),
    defer_named_return_with_branch_runtime => ("package main; import \"fmt\"; func build(flag bool) (result int) { defer func() { result++ }(); if flag { return 5 }; return 2 }; func main() { fmt.Println(build(true)); }", vec!["6"]),
    defer_multiple_prints_runtime => ("package main; import \"fmt\"; func main() { defer fmt.Println(\"a\"); defer fmt.Println(\"b\"); fmt.Println(\"c\"); }", vec!["c", "b", "a"]),
    defer_with_argument_copy_runtime => ("package main; import \"fmt\"; func show(v int) { fmt.Println(v) }; func main() { value := 2; defer show(value); value = 9; }", vec!["2"]),
    recover_from_integer_panic_runtime => ("package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover()) }(); panic(12) }; func main() { run() }", vec!["12"]),
    defer_closure_reads_latest_outer_runtime => ("package main; import \"fmt\"; func main() { value := 1; defer func() { fmt.Println(value) }(); value = 3; }", vec!["3"]),
}

compile_cases! {
    recover_returns_value_compile => "package main; import \"fmt\"; func run() { defer func() { fmt.Println(recover()) }(); panic(\"boom\") }; func main() { run() }",
    defer_with_named_return_compile => "package main; func build() (result int) { defer func() { result++ }(); return 1 }; func main() { _ = build }",
    defer_with_pointer_param_compile => "package main; func main() { value := 1; defer func(ptr *int) { *ptr = 2 }(&value) }",
    recover_in_deferred_func_compile => "package main; func main() { defer func() { _ = recover() }(); panic(\"x\") }",
    panic_with_string_compile => "package main; func main() { panic(\"boom\") }",
    panic_with_int_compile => "package main; func main() { panic(1) }",
    nested_defer_compile => "package main; func main() { defer func() { defer func() {}() }() }",
    defer_in_if_compile => "package main; func main() { if true { defer func() {}() } }",
    defer_in_switch_compile => "package main; func main() { switch 1 { case 1: defer func() {}() } }",
    defer_in_loop_compile => "package main; func main() { for i := 0; i < 2; i++ { defer func() { _ = i }() } }",
    recover_without_panic_compile => "package main; func main() { _ = recover() }",
    defer_call_named_function_compile => "package main; func cleanup() {}; func main() { defer cleanup() }",
    defer_method_call_compile => "package main; type counter struct{}; func (counter) clean() {}; func main() { value := counter{}; defer value.clean() }",
    defer_pointer_method_call_compile => "package main; type counter struct{}; func (c *counter) clean() {}; func main() { value := &counter{}; defer value.clean() }",
    defer_with_multiple_returns_compile => "package main; func build() int { defer func() {}(); if true { return 1 }; return 2 }; func main() { _ = build }",
    recover_in_nested_defer_compile => "package main; func main() { defer func() { defer func() { _ = recover() }() }(); panic(\"x\") }",
    panic_after_defer_compile => "package main; func main() { defer func() {}(); panic(\"x\") }",
    defer_modify_named_result_compile => "package main; func build() (result int) { defer func() { result += 2 }(); return 1 }; func main() { _ = build }",
    defer_modify_slice_compile => "package main; func main() { values := []int{1}; defer func() { values[0] = 2 }() }",
    defer_modify_map_compile => "package main; func main() { values := map[string]int{\"a\": 1}; defer func() { values[\"a\"] = 2 }() }",
    defer_closure_capture_compile => "package main; func main() { value := 1; defer func() { _ = value }() }",
    recover_assigned_compile => "package main; func main() { defer func() { value := recover(); _ = value }(); panic(\"x\") }",
    panic_in_helper_compile => "package main; func boom() { panic(\"x\") }; func main() { defer func() { _ = recover() }(); boom() }",
    defer_inside_anonymous_function_compile => "package main; func main() { func() { defer func() {}() }() }",
    defer_in_returned_function_compile => "package main; func build() func() { return func() { defer func() {}() } }; func main() { _ = build() }",
    recover_value_compare_compile => "package main; func main() { defer func() { value := recover(); _ = (value != nil) }(); panic(1) }",
}
