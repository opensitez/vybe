// vybe-test: go/slices_sort_equal_extended/slices_replace_full_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{1, 2, 3}
_ = slices.Replace(s, 0, 3, 9, 8) }
