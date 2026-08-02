// vybe-test: go/interface_assertion_extended/assert_concrete_int_to_string_compile_fail
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile-fail

package main
func main() { var x int = 1
_ = x.(string) }
