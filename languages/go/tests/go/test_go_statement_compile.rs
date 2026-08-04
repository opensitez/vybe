//! Compile coverage for `go` statement forms: closures, method calls, parameterized
//! anonymous functions, and loop-driven goroutine spawning.

compile_cases! {
    // Closure capture
    go_closure_capture_local_int_compile => "package main; func main() { x := 1; go func() { _ = x }() }",
    go_closure_capture_local_string_compile => "package main; func main() { msg := \"hi\"; go func() { _ = msg }() }",
    go_closure_capture_struct_field_compile => "package main; type box struct { n int }; func main() { b := box{n: 2}; go func() { _ = b.n }() }",
    go_closure_capture_pointer_target_compile => "package main; func main() { n := 3; p := &n; go func() { _ = *p }() }",
    go_closure_capture_slice_element_compile => "package main; func main() { s := []int{1, 2}; go func() { _ = s[0] }() }",
    go_closure_capture_outer_parameter_compile => "package main; func worker(x int) { go func() { _ = x }() }; func main() { worker(4) }",
    go_closure_capture_multiple_locals_compile => "package main; func main() { a, b := 1, 2; go func() { _, _ = a, b }() }",
    go_closure_with_defer_inside_compile => "package main; func main() { go func() { defer func() { _ = 1 }(); _ = 0 }() }",
    go_closure_nested_spawn_compile => "package main; func main() { go func() { go func() { _ = 1 }() }() }",
    go_closure_mutate_captured_variable_compile => "package main; func main() { n := 0; go func() { n = 5 }() }",

    // Method and function-value goroutine targets
    go_pointer_receiver_method_compile => "package main; type worker struct { n int }; func (w *worker) bump() { w.n++ }; func main() { w := &worker{}; go w.bump() }",
    go_method_on_interface_value_compile => "package main; type runner interface { run() }; type worker struct{}; func (worker) run() {}; func main() { var r runner = worker{}; go r.run() }",
    go_promoted_embedded_method_compile => "package main; type inner struct{}; func (inner) work() {}; type outer struct { inner }; func main() { o := outer{}; go o.work() }",
    go_method_on_named_type_compile => "package main; type counter int; func (c counter) inc() {}; func main() { var c counter; go c.inc() }",
    go_method_on_struct_literal_compile => "package main; type worker struct{}; func (worker) run() {}; func main() { go worker{}.run() }",
    go_method_value_as_goroutine_compile => "package main; type worker struct{}; func (worker) run() {}; func main() { w := worker{}; fn := w.run; go fn() }",
    go_named_package_function_compile => "package main; func tick() {}; func main() { go tick() }",

    // Anonymous function with arguments
    go_anon_func_one_int_arg_compile => "package main; func main() { go func(v int) { _ = v }(7) }",
    go_anon_func_two_args_compile => "package main; func main() { go func(a int, b string) { _, _ = a, b }(1, \"x\") }",
    go_anon_func_string_arg_compile => "package main; func main() { go func(label string) { _ = label }(\"go\") }",
    go_anon_func_variadic_args_compile => "package main; func main() { go func(nums ...int) { _ = len(nums) }(1, 2, 3) }",
    go_anon_func_param_and_closure_capture_compile => "package main; func main() { base := 10; go func(delta int) { _ = base + delta }(3) }",

    // Loop-driven spawning
    go_spawn_for_loop_index_capture_compile => "package main; func main() { for i := 0; i < 3; i++ { go func() { _ = i }() } }",
    go_spawn_for_loop_pass_index_arg_compile => "package main; func main() { for i := 0; i < 3; i++ { go func(idx int) { _ = idx }(i) } }",
    go_spawn_range_over_slice_compile => "package main; func main() { for _, v := range []int{1, 2} { go func() { _ = v }() } }",
    go_spawn_range_pass_value_arg_compile => "package main; func main() { for _, v := range []int{4, 5} { go func(n int) { _ = n }(v) } }",
    go_spawn_nested_for_loops_compile => "package main; func main() { for i := 0; i < 2; i++ { for j := 0; j < 2; j++ { go func() { _, _ = i, j }() } } }",
}
