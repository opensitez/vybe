// vybe-test: go/init_function_order/init_calls_two_package_helpers_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var ready bool
func arm() { ready = true }
func confirm() { _ = ready }
func init() { arm() }
func init() { confirm() }
func main() {}
