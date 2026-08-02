// vybe-test: go/blank_identifier_extended/blank_discard_composite_array_literal_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { _ = [2]int{1, 2}
a := [2]int{1}
_ = a[0] }
