// vybe-test: go/init_function_order/init_grouped_var_block_used_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var ( x int; y int )
func init() { x = 1
y = 2 }
func init() { x = x + y }
func main() { _ = x }
