// vybe-test: go/slices_sort_equal_extended/slices_sort_func_rune_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []rune{'日', 'a', 'b'}
slices.SortFunc(s, func(a, b rune) int { if a < b { return -1 }; if a > b { return 1 }; return 0 }) }
