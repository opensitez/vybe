// vybe-test: go/type_conversions_extra/conversion_in_return_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func cast(v int) float64 { return float64(v) }
func main() { _ = cast }
