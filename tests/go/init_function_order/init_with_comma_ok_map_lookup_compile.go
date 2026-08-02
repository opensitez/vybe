// vybe-test: go/init_function_order/init_with_comma_ok_map_lookup_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var m = map[string]int{"a": 1}
var v int
func init() { v, _ = m["a"] }
func main() { _ = v }
