// vybe-test: go/interface_assertion_extended/comma_ok_on_non_interface_compile_fail
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile-fail

package main
func main() { x := 1
_, ok := x.(int)
_ = ok }
