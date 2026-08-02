// vybe-test: go/higher_order_functions/function_type_in_struct
// origin: languages/go/tests/go/test_higher_order_functions.rs
// vybe-test-mode: compile

package main
type handler struct { fn func() }
func main() { _ = handler{} }
