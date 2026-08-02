// vybe-test: go/init_function_order/init_method_on_named_type_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
type counter int
func (c *counter) inc() { *c++ }
var total counter
func init() { total.inc() }
func main() { _ = total }
