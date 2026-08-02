// vybe-test: go/init_function_order/init_reads_package_var_from_prior_init_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var a int
var b int
func init() { a = 7 }
func init() { b = a + 1 }
func main() { _ = b }
