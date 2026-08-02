// vybe-test: go/nil_zero_semantics_extra/zero_value_struct_literal_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func main() { _ = counter{} }
