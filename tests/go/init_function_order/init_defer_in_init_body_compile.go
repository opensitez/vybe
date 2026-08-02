// vybe-test: go/init_function_order/init_defer_in_init_body_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var x int
func init() { defer func() { x++ }()
x = 1 }
func main() { _ = x }
