// vybe-test: go/interface_assertion_extended/assert_to_any_alias_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
func main() { var v any = 1
_, ok := v.(int)
_ = ok }
