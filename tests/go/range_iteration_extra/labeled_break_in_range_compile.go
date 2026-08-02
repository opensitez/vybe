// vybe-test: go/range_iteration_extra/labeled_break_in_range_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { outer: for _, row := range [][]int{{1}} { for _, value := range row { _ = value
break outer } } }
