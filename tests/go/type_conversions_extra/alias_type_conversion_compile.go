// vybe-test: go/type_conversions_extra/alias_type_conversion_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type count = int
func main() { _ = count(2) }
