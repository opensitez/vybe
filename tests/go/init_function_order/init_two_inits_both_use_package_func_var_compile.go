// vybe-test: go/init_function_order/init_two_inits_both_use_package_func_var_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var apply func(int) int
func init() { apply = func(n int) int { return n + 1 } }
func init() { _ = apply(2) }
func main() {}
