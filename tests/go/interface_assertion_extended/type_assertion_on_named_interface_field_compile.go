// vybe-test: go/interface_assertion_extended/type_assertion_on_named_interface_field_compile
// origin: languages/go/tests/go/test_interface_assertion_extended.rs
// vybe-test-mode: compile

package main
type speaker interface { talk() string }
type holder struct { value interface{} }
func main() { h := holder{value: 1}
_, _ = h.value.(int) }
