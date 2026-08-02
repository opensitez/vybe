// vybe-test: go/slices_sort_equal_extended/slices_is_sorted_func_float64
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []float64{1.0, 2.0}
_ = slices.IsSortedFunc(s, func(a, b float64) int { if a < b { return -1 }; if a > b { return 1 }; return 0 }) }
