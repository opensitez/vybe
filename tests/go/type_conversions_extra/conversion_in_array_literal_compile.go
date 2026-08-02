// vybe-test: go/type_conversions_extra/conversion_in_array_literal_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = [2]int{int(1.2), int(2.3)} }
