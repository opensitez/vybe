// vybe-test: go/slices_sort_equal_extended/slices_sort_stable_func_struct_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
type Item struct { N int }
func main() { s := []Item{{2}, {1}, {2}}
slices.SortStableFunc(s, func(a, b Item) int { if a.N < b.N { return -1 }; if a.N > b.N { return 1 }; return 0 }) }
