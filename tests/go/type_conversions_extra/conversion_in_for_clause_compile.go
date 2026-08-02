// vybe-test: go/type_conversions_extra/conversion_in_for_clause_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { for i := int(0); i < 1; i++ { _ = i } }
