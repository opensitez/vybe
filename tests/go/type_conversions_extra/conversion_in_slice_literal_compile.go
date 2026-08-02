// vybe-test: go/type_conversions_extra/conversion_in_slice_literal_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = []float64{float64(1), float64(2)} }
