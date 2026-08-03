//! Function literals and closures: capture, assignment, IIFE, returned closures,
//! http.HandlerFunc, recursive var bindings. Distinct from `test_closures.rs`
//! (basic closure forms) and `test_higher_order_functions.rs` (HOF patterns).

go_run_cases! {
    literal_capture_outer_int_increment =>
        ("package main; import \"fmt\"; func main() { n := 0; inc := func() { n++ }; inc(); inc(); fmt.Println(n) }", vec!["2"]),
    literal_capture_outer_string_mutate =>
        ("package main; import \"fmt\"; func main() { s := \"a\"; appendChar := func(c string) { s += c }; appendChar(\"b\"); fmt.Println(s) }", vec!["ab"]),
    literal_capture_multiple_outer_vars =>
        ("package main; import \"fmt\"; func main() { a := 2; b := 3; sum := func() int { return a + b }; fmt.Println(sum()) }", vec!["5"]),
    literal_assign_to_typed_variable =>
        ("package main; import \"fmt\"; func main() { var fn func(int) int = func(x int) int { return x * 3 }; fmt.Println(fn(4)) }", vec!["12"]),
    literal_assign_to_short_var =>
        ("package main; import \"fmt\"; func main() { double := func(x int) int { return x * 2 }; fmt.Println(double(7)) }", vec!["14"]),
    literal_call_immediately_iife =>
        ("package main; import \"fmt\"; func main() { result := func(a int, b int) int { return a + b }(10, 20); fmt.Println(result) }", vec!["30"]),
    literal_iife_no_args =>
        ("package main; import \"fmt\"; func main() { func() { fmt.Println(\"run\") }() }", vec!["run"]),
    literal_iife_returns_value =>
        ("package main; import \"fmt\"; func main() { n := func() int { return 99 }(); fmt.Println(n) }", vec!["99"]),
    return_closure_from_function =>
        ("package main; import \"fmt\"; func makeAdder(base int) func(int) int { return func(x int) int { return base + x } }; func main() { add5 := makeAdder(5); fmt.Println(add5(3)) }", vec!["8"]),
    return_closure_captures_parameter =>
        ("package main; import \"fmt\"; func scale(factor int) func(int) int { return func(v int) int { return v * factor } }; func main() { triple := scale(3); fmt.Println(triple(4)) }", vec!["12"]),
    closure_passed_as_argument =>
        ("package main; import \"fmt\"; func apply(fn func(int) int, x int) int { return fn(x) }; func main() { sq := func(n int) int { return n * n }; fmt.Println(apply(sq, 6)) }", vec!["36"]),
    closure_stored_in_struct_field =>
        ("package main; import \"fmt\"; type holder struct { fn func() int }; func main() { h := holder{fn: func() int { return 42 }}; fmt.Println(h.fn()) }", vec!["42"]),
    closure_stored_in_map_value =>
        ("package main; import \"fmt\"; func main() { ops := map[string]func(int, int) int{ \"add\": func(a, b int) int { return a + b } }; fmt.Println(ops[\"add\"](3, 4)) }", vec!["7"]),
    closure_stored_in_slice =>
        ("package main; import \"fmt\"; func main() { fns := []func(int) int{ func(x int) int { return x + 1 }, func(x int) int { return x + 2 } }; fmt.Println(fns[0](5)); fmt.Println(fns[1](5)) }", vec!["6", "7"]),
    recursive_closure_via_var_binding =>
        ("package main; import \"fmt\"; func main() { var fib func(int) int; fib = func(n int) int { if n < 2 { return n }; return fib(n-1) + fib(n-2) }; fmt.Println(fib(6)) }", vec!["8"]),
    recursive_closure_factorial =>
        ("package main; import \"fmt\"; func main() { var fact func(int) int; fact = func(n int) int { if n <= 1 { return 1 }; return n * fact(n-1) }; fmt.Println(fact(5)) }", vec!["120"]),
    closure_capture_loop_var_with_param =>
        ("package main; import \"fmt\"; func main() { sum := 0; for i := 1; i <= 3; i++ { func(n int) { sum += n }(i) }; fmt.Println(sum) }", vec!["6"]),
    closure_modifies_outer_slice =>
        ("package main; import \"fmt\"; func main() { items := []int{}; push := func(v int) { items = append(items, v) }; push(1); push(2); fmt.Println(len(items)); fmt.Println(items[1]) }", vec!["2", "2"]),
    closure_returns_closure =>
        ("package main; import \"fmt\"; func outer(x int) func(int) func(int) int { return func(y int) func(int) int { return func(z int) int { return x + y + z } } }; func main() { fn := outer(1)(2); fmt.Println(fn(3)) }", vec!["6"]),
    closure_with_defer_inside =>
        ("package main; import \"fmt\"; func main() { run := func() { defer fmt.Println(\"done\"); fmt.Println(\"go\") }; run() }", vec!["go", "done"]),
    closure_as_named_return_helper =>
        ("package main; import \"fmt\"; func compute() int { transform := func(x int) int { return x + 10 }; return transform(5) }; func main() { fmt.Println(compute()) }", vec!["15"]),
    closure_capture_bool_toggle =>
        ("package main; import \"fmt\"; func main() { on := false; flip := func() { on = !on }; flip(); flip(); fmt.Println(on) }", vec!["false"]),
    closure_capture_struct_field =>
        ("package main; import \"fmt\"; type counter struct { n int }; func main() { c := counter{n: 0}; bump := func() { c.n++ }; bump(); bump(); fmt.Println(c.n) }", vec!["2"]),
    closure_with_named_result_params =>
        ("package main; import \"fmt\"; func main() { divide := func(a, b int) (q int, r int) { q = a / b; r = a % b; return }; q, r := divide(10, 3); fmt.Println(q); fmt.Println(r) }", vec!["3", "1"]),
    closure_nil_check_before_call =>
        ("package main; import \"fmt\"; func main() { var fn func() = nil; if fn != nil { fn() } else { fmt.Println(\"nil\") } }", vec!["nil"]),
    closure_reassign_variable =>
        ("package main; import \"fmt\"; func main() { fn := func() int { return 1 }; fn = func() int { return 2 }; fmt.Println(fn()) }", vec!["2"]),
    closure_compare_equality =>
        ("package main; import \"fmt\"; func main() { a := func() {}; b := a; fmt.Println(a == nil); fmt.Println(b == nil) }", vec!["false", "false"]),
    two_closures_share_outer_state =>
        ("package main; import \"fmt\"; func main() { n := 0; inc := func() { n++ }; get := func() int { return n }; inc(); inc(); fmt.Println(get()) }", vec!["2"]),
    closure_in_select_case =>
        ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 1; select { case fn := func(v int) int { return v }(<-ch): fmt.Println(fn) } }", vec!["1"]),
    closure_returned_from_if_branch =>
        ("package main; import \"fmt\"; func pick(positive bool) func(int) int { if positive { return func(x int) int { return x } }; return func(x int) int { return -x } }; func main() { fmt.Println(pick(false)(5)) }", vec!["-5"]),
    closure_with_variadic_params =>
        ("package main; import \"fmt\"; func main() { sum := func(nums ...int) int { total := 0; for _, n := range nums { total += n }; return total }; fmt.Println(sum(1, 2, 3)) }", vec!["6"]),
    closure_capture_array =>
        ("package main; import \"fmt\"; func main() { arr := [2]int{1, 2}; read := func(i int) int { return arr[i] }; fmt.Println(read(1)) }", vec!["2"]),
    closure_capture_map =>
        ("package main; import \"fmt\"; func main() { m := map[string]int{\"a\": 1}; lookup := func(k string) int { return m[k] }; fmt.Println(lookup(\"a\")) }", vec!["1"]),
    closure_as_interface_method =>
        ("package main; import \"fmt\"; type runner interface { run() int }; func main() { var r runner = runnerFunc(func() int { return 7 }); fmt.Println(r.run()) }; type runnerFunc func() int; func (f runnerFunc) run() int { return f() }", vec!["7"]),
    nested_closure_three_levels =>
        ("package main; import \"fmt\"; func main() { level1 := func(a int) func(b int) func(c int) int { return func(b int) func(c int) int { return func(c int) int { return a + b + c } } }; fn := level1(1)(2); fmt.Println(fn(3)) }", vec!["6"]),
    closure_with_panic_recover =>
        ("package main; import \"fmt\"; func main() { safe := func() { defer func() { fmt.Println(recover() != nil) }(); panic(\"x\") }; safe() }", vec!["true"]),
    closure_filter_slice =>
        ("package main; import \"fmt\"; func main() { nums := []int{1, 2, 3, 4}; evens := func() []int { out := []int{}; for _, n := range nums { if n%2 == 0 { out = append(out, n) } }; return out }; r := evens(); fmt.Println(len(r)); fmt.Println(r[0]) }", vec!["2", "2"]),
    closure_string_builder_pattern =>
        ("package main; import \"fmt\"; func main() { parts := []string{\"go\", \"lang\"}; join := func(sep string) string { s := \"\"; for i, p := range parts { if i > 0 { s += sep }; s += p }; return s }; fmt.Println(join(\"-\")) }", vec!["go-lang"]),
    closure_mutual_via_vars =>
        ("package main; import \"fmt\"; func main() { var even func(int) bool; var odd func(int) bool; even = func(n int) bool { if n == 0 { return true }; return odd(n-1) }; odd = func(n int) bool { if n == 0 { return false }; return even(n-1) }; fmt.Println(even(4)); fmt.Println(odd(3)) }", vec!["true", "true"]),
    closure_capture_pointer =>
        ("package main; import \"fmt\"; func main() { n := 5; ptr := &n; bump := func() { *ptr = *ptr + 1 }; bump(); fmt.Println(n) }", vec!["6"]),
    closure_in_for_range =>
        ("package main; import \"fmt\"; func main() { sum := 0; for _, v := range []int{1, 2, 3} { func(x int) { sum += x }(v) }; fmt.Println(sum) }", vec!["6"]),
    closure_return_same_capture =>
        ("package main; import \"fmt\"; func main() { base := 10; mk := func() func() int { return func() int { return base } }; fmt.Println(mk()()) }", vec!["10"]) }

go_compile_cases! {
    closure_as_http_handler_func_compile =>
        "package main; import \"net/http\"; func main() { var h http.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusOK) }) }",
    closure_handler_func_serve_http_compile =>
        "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _, _ = w.Write([]byte(\"ok\")) }).ServeHTTP(nil, nil) }",
    closure_handler_func_read_header_compile =>
        "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _ = r.Header.Get(\"Accept\") }) }",
    closure_handler_func_with_capture_compile =>
        "package main; import \"net/http\"; func main() { prefix := \"x\"; http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _ = prefix + r.Method }) }",
    closure_assigned_to_interface_var_compile =>
        "package main; func main() { var fn func(int) int = func(x int) int { return x }; _ = fn(1) }",
    closure_returned_from_named_func_compile =>
        "package main; func mk() func() { return func() {} }; func main() { _ = mk() }",
    recursive_closure_var_decl_compile =>
        "package main; func main() { var f func(int) int; f = func(n int) int { if n == 0 { return 0 }; return f(n-1) }; _ = f(3) }",
    closure_in_struct_literal_compile =>
        "package main; type pair struct { fn func() }; func main() { _ = pair{fn: func() {}} }",
    closure_as_map_value_compile =>
        "package main; func main() { m := map[int]func(){ 1: func() {} }; m[1]() }",
    closure_in_goroutine_arg_compile =>
        "package main; func main() { go func() {}() }",
    closure_capture_outer_in_defer_compile =>
        "package main; func main() { n := 1; defer func() { _ = n }() }",
    closure_with_blank_import_capture_compile =>
        "package main; import \"fmt\"; func main() { fn := func() { fmt.Println(1) }; fn() }",
    iife_in_expression_compile =>
        "package main; func main() { _ = func(x int) int { return x }(2) }",
    closure_type_in_slice_compile =>
        "package main; func main() { fns := []func(){ func() {}, func() {} }; _ = fns }",
    closure_passed_to_higher_order_compile =>
        "package main; func call(fn func(int) int, x int) int { return fn(x) }; func main() { _ = call(func(x int) int { return x }, 1) }",
    closure_with_named_return_compile =>
        "package main; func main() { fn := func() (n int) { n = 1; return }; _, _ = fn(), fn() }",
    mutual_recursive_closure_compile =>
        "package main; var a func(int) bool; var b func(int) bool; func init() { a = func(n int) bool { return b(n-1) }; b = func(n int) bool { return a(n-1) } }; func main() { _ = a(2) }",
    closure_handler_func_status_code_compile =>
        "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusNotFound) }) }",
    closure_in_select_compile =>
        "package main; func main() { ch := make(chan int); select { case <-func() chan int { return ch }(): } }" }
