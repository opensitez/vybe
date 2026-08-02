// vybe-test: go/range_over_int/range_int_labeled_continue_outer_compile
// origin: languages/go/tests/go/test_range_over_int.rs
// vybe-test-mode: compile

package main
func main() { outer: for i := range 3 { if i == 1 { continue outer }
_ = i } }
