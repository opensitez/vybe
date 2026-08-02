// vybe-test: go/sort_slice_find/sort_find_not_found
// origin: languages/go/tests/go/test_sort_slice_find.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { s := []int{1,3,5}
_, _ = sort.Find(s, 0, func(i int) int { return s[i] }) }
