// vybe-test: go/interface_assertion_extended/type_assertion_to_interface_type_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
type reader interface { read() int }
func main() { var v interface{} = 1
_, ok := v.(reader)
_ = ok }
