//! Functions, returns, and call semantics — one distinct rule per test.

go_run_cases! {
    multi_return_swap => ("package main; import \"fmt\"; func pair() (int, string) { return 1, \"a\" }; func main() { a, b := pair(); fmt.Println(a, b) }", vec!["1 a"]),
    multi_return_ignore_with_blank => ("package main; import \"fmt\"; func pair() (int, string) { return 2, \"b\" }; func main() { _, s := pair(); fmt.Println(s) }", vec!["b"]),
    named_result_initialized_zero => ("package main; import \"fmt\"; func f() (n int, s string) { return }; func main() { n, s := f(); fmt.Println(n, s == \"\") }", vec!["0 true"]),
    named_result_assignment_before_bare_return => ("package main; import \"fmt\"; func f() (n int) { n = 9; return }; func main() { fmt.Println(f()) }", vec!["9"]),
    variadic_empty_call => ("package main; import \"fmt\"; func sum(xs ...int) int { t := 0; for _, x := range xs { t += x }; return t }; func main() { fmt.Println(sum()) }", vec!["0"]),
    variadic_slice_spread => ("package main; import \"fmt\"; func sum(xs ...int) int { t := 0; for _, x := range xs { t += x }; return t }; func main() { xs := []int{1,2}; fmt.Println(sum(xs...)) }", vec!["3"]),
    recursive_base_case => ("package main; import \"fmt\"; func fact(n int) int { if n <= 1 { return 1 }; return n * fact(n-1) }; func main() { fmt.Println(fact(4)) }", vec!["24"]),
    closure_mutates_outer => ("package main; import \"fmt\"; func main() { n := 0; f := func() { n++ }; f(); f(); fmt.Println(n) }", vec!["2"]),
    defer_modifies_named_result => ("package main; import \"fmt\"; func f() (n int) { defer func() { n++ }(); return 1 }; func main() { fmt.Println(f()) }", vec!["2"]),
    function_as_value_nil_compare => ("package main; import \"fmt\"; func main() { var f func(); fmt.Println(f == nil) }", vec!["true"]),
    higher_order_map => ("package main; import \"fmt\"; func mapInts(xs []int, f func(int) int) []int { out := make([]int, len(xs)); for i, v := range xs { out[i] = f(v) }; return out }; func main() { fmt.Println(mapInts([]int{1,2}, func(x int) int { return x*2 })[1]) }", vec!["4"]),
    method_call_passed_as_value => ("package main; import \"fmt\"; type S struct{}; func (S) ID() int { return 7 }; func main() { var s S; fmt.Println(s.ID()) }", vec!["7"]),
    init_before_main_order => ("package main; import \"fmt\"; var n = func() int { return 3 }(); func main() { fmt.Println(n) }", vec!["3"]),
    package_level_func_forward_ref => ("package main; import \"fmt\"; func main() { fmt.Println(g()) }; func g() int { return 5 }", vec!["5"]),
}

go_compile_cases! {
    return_discard_values_compile => "package main; func f() (int, int) { return 1, 2 }; func main() { f() }",
    naked_return_in_nested_func_compile => "package main; func f() (n int) { n = 1; return }; func main() { _ = f() }",
    func_literal_type_inference => "package main; func main() { _ = func(x int) int { return x } }",
    call_nil_func_compile => "package main; func main() { var f func(); f() }",
    defer_call_args_evaluated_early => "package main; func main() { defer func(int) {}(func() int { return 1 }()) }",
    method_expression_with_instantiation => "package main; type T int; func (T) M() {}; func main() { _ = T.M }",
}
