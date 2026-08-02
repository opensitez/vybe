// vybe-test: go/slices_sort_equal_extended/slices_sort_func_three_way_tie
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{2, 2, 2}
slices.SortFunc(s, func(a, b int) int { return 0 }) }
