// vybe-test: go/generics_constraints_extended/generic_ordered_heap_not_needed_sortfunc
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
import "cmp"
import "slices"
func SortDesc[T cmp.Ordered](s []T) { slices.SortFunc(s, func(a, b T) int { if a > b { return -1 }; if a < b { return 1 }; return 0 }) }
func main() { data := []int{1, 3, 2}
SortDesc(data) }
