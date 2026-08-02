// vybe-test: go/interface_assertion_extended/type_assertion_comma_ok_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
func main() { var v interface{} = 1
_, ok := v.(int)
_ = ok }
