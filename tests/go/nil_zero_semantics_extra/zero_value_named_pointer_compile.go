// vybe-test: go/nil_zero_semantics_extra/zero_value_named_pointer_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
type score int
func main() { var value *score
_ = value }
