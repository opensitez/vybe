// vybe-test: go/type_conversions_extra/conversion_in_struct_literal_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type holder struct { value float64 }
func main() { _ = holder{value: float64(4)} }
