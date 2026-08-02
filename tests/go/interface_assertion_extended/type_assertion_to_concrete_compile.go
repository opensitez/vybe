// vybe-test: go/interface_assertion_extended/type_assertion_to_concrete_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
func main() { var v interface{} = 1
_ = v.(int) }
