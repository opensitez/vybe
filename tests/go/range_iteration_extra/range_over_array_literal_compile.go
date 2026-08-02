// vybe-test: go/range_iteration_extra/range_over_array_literal_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { for _, value := range [2]int{1, 2} { _ = value } }
