// vybe-test: go/init_function_order/init_with_if_init_short_decl_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var ok bool
func init() { if n := 3; n > 0 { ok = true } }
func main() { _ = ok }
