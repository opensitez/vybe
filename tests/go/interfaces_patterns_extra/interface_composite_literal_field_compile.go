// vybe-test: go/interfaces_patterns_extra/interface_composite_literal_field_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type holder struct { value interface{} }
func main() { _ = holder{value: 1} }
