// vybe-test: go/blank_identifier_extended/blank_multi_assign_swap_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { a, b := 1, 2
a, b = b, a
_, _ = a, b }
