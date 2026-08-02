// vybe-test: go/init_function_order/init_with_type_switch_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var tag string
func init() { var v interface{} = 3
switch v.(type) { case int: tag = "int" default: tag = "other" } }
func main() { _ = tag }
