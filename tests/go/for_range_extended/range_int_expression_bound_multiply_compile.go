// vybe-test: go/for_range_extended/range_int_expression_bound_multiply_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 * 2 { _ = i } }
