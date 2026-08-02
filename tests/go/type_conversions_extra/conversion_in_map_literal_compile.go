// vybe-test: go/type_conversions_extra/conversion_in_map_literal_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = map[string]float64{"a": float64(3)} }
