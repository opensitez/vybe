// vybe-test: go/nil_zero_semantics_extra/blank_identifier_zero_values_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var a int
var b string
_, _ = a, b }
