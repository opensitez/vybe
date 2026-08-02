// vybe-test: go/init_function_order/init_assigns_func_value_then_calls_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var double func(int) int
func init() { double = func(n int) int { return n * 2 } }
func init() { _ = double(3) }
func main() {}
