// vybe-test: go/init_function_order/init_with_range_loop_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var sum int
func init() { for i := range 4 { sum += i } }
func main() { _ = sum }
