// vybe-test: go/slices_sort_equal_extended/slices_binary_search_func_empty
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { _, _ = slices.BinarySearchFunc([]int{}, 1, func(a, b int) int { return a - b }) }
