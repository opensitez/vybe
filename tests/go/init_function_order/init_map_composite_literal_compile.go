// vybe-test: go/init_function_order/init_map_composite_literal_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var table map[string]int
func init() { table = map[string]int{"k": 9} }
func main() { _ = table["k"] }
