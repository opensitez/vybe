// vybe-test: go/type_conversions_extra/conversion_of_struct_field_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type holder struct { value int }
func main() { h := holder{value: 8}
_ = float64(h.value) }
