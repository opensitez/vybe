// vybe-test: go/interface_assertion_extended/type_assertion_on_non_interface_expr_compile_fail
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile-fail

package main
func main() { x := 1
_ = x.(int) }
