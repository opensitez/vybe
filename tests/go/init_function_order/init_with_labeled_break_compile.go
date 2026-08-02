// vybe-test: go/init_function_order/init_with_labeled_break_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var count int
func init() { outer: for i := 0; i < 5; i++ { count++
break outer } }
func main() { _ = count }
