// vybe-test: go/slices_sort_equal_extended/slices_compare_empty_vs_nil
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { var a []int
_ = slices.Compare(a, []int{}) }
