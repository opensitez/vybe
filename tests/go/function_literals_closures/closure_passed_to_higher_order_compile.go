// vybe-test: go/function_literals_closures/closure_passed_to_higher_order_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func call(fn func(int) int, x int) int { return fn(x) }
func main() { _ = call(func(x int) int { return x }, 1) }
