// vybe-test: go/interface_assertion_extended/type_assertion_result_used_in_switch_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
func main() { var v interface{} = "x"
switch s := v.(type) { case string: _ = s } }
