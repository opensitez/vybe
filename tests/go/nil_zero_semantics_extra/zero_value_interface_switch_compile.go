// vybe-test: go/nil_zero_semantics_extra/zero_value_interface_switch_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var value interface{}
switch value.(type) { default: } }
