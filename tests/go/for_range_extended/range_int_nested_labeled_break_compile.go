// vybe-test: go/for_range_extended/range_int_nested_labeled_break_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { outer: for i := range 2 { for j := range 2 { if j == 1 { break outer }
_ = i } } }
