// vybe-test: go/init_function_order/init_uses_const_arithmetic_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
const ( A = 2; B = A * 3 )
var c int
func init() { c = B + 1 }
func main() { _ = c }
