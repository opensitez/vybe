// vybe-test: go/slices_sort_equal_extended/slices_is_sorted_uint_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { _ = slices.IsSorted([]uint{1, 2, 3}) }
