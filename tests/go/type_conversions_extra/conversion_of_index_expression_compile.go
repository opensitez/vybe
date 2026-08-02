// vybe-test: go/type_conversions_extra/conversion_of_index_expression_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{6}
_ = float64(values[0]) }
