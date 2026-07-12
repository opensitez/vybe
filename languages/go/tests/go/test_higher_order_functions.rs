//! Higher-order functions: functions as values, returning closures, callbacks.

go_run_cases! {
    function_value_stored_and_called => ("package main; import \"fmt\"; func main() { var f func(int) int = func(x int) int { return x + 1 }; fmt.Println(f(4)) }", vec!["5"]),
    function_returns_closure => ("package main; import \"fmt\"; func adder(base int) func(int) int { return func(x int) int { return base + x } }; func main() { inc := adder(10); fmt.Println(inc(3)) }", vec!["13"]),
    callback_reduces_slice => ("package main; import \"fmt\"; func fold(nums []int, combine func(int,int) int, init int) int { acc := init; for _, n := range nums { acc = combine(acc, n) }; return acc }; func main() { fmt.Println(fold([]int{1,2,3}, func(a,b int) int { return a+b }, 0)) }", vec!["6"]),
    filter_with_predicate => ("package main; import \"fmt\"; func keep(nums []int, ok func(int) bool) []int { out := []int{}; for _, n := range nums { if ok(n) { out = append(out, n) } }; return out }; func main() { r := keep([]int{1,2,3,4}, func(n int) bool { return n%2 == 0 }); fmt.Println(len(r)); fmt.Println(r[0]) }", vec!["2", "2"]),
    compose_two_functions => ("package main; import \"fmt\"; func compose(f func(int) int, g func(int) int) func(int) int { return func(x int) int { return f(g(x)) } }; func main() { double := func(x int) int { return x * 2 }; inc := func(x int) int { return x + 1 }; h := compose(double, inc); fmt.Println(h(3)) }", vec!["8"]),
}

go_compile_cases! {
    higher_order_generic_callback => "package main; func mapInts(src []int, f func(int) int) []int { out := make([]int, len(src)); for i, v := range src { out[i] = f(v) }; return out }; func main() { _ = mapInts([]int{1}, func(x int) int { return x }) }",
    function_type_in_struct => "package main; type handler struct { fn func() }; func main() { _ = handler{} }",
    mutual_recursive_functions => "package main; var even func(int) bool; var odd func(int) bool; func init() { even = func(n int) bool { if n == 0 { return true }; return odd(n-1) }; odd = func(n int) bool { if n == 0 { return false }; return even(n-1) } }; func main() { _ = even(4) }",
}
