// vybe-test: go/interface_assertion_extended/type_assertion_to_pointer_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
type node struct{}
func main() { var v interface{} = &node{}
_ = v.(*node) }
