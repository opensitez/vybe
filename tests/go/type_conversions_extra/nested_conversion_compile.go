// vybe-test: go/type_conversions_extra/nested_conversion_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = int(float64(7)) }
