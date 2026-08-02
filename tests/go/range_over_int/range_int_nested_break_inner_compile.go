// vybe-test: go/range_over_int/range_int_nested_break_inner_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 { for j := range 2 { if j == 1 { break }
_ = j } } }
