// vybe-test: go/init_function_order/init_for_range_over_string_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var count int
func init() { for range "go" { count++ } }
func main() { _ = count }
