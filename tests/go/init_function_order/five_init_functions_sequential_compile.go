// vybe-test: go/init_function_order/five_init_functions_sequential_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var phase int
func init() { phase = 1 }
func init() { phase++ }
func init() { phase++ }
func init() { phase++ }
func init() { phase++ }
func main() { _ = phase }
