// vybe-test: go/nil_zero_semantics_extra/zero_value_pointer_deref_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var value *int
_ = *value }
