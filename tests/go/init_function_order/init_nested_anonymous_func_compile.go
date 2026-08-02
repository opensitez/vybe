// vybe-test: go/init_function_order/init_nested_anonymous_func_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var n int
func init() { func() { n = 4 }() }
func main() { _ = n }
