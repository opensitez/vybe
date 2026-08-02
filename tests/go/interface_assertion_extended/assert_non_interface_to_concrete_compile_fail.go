// vybe-test: go/interface_assertion_extended/assert_non_interface_to_concrete_compile_fail
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile-fail

package main
type a struct{}
func main() { var x a
_ = x.(a) }
