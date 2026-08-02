// vybe-test: go/interface_assertion_extended/interface_assertion_in_assignment_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
type fmtStringer interface { String() string }
func main() { var v interface{} = 1
_, ok := v.(fmtStringer)
_ = ok }
