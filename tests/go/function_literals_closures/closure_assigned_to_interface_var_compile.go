// vybe-test: go/function_literals_closures/closure_assigned_to_interface_var_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { var fn func(int) int = func(x int) int { return x }
_ = fn(1) }
