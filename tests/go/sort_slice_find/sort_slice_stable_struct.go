// vybe-test: go/sort_slice_find/sort_slice_stable_struct
// origin: languages/go/tests/go/test_sort_slice_find.rs
// vybe-test-mode: compile

package main
import "sort"
type P struct { k int }
func main() { s := []P{{1},{2}}
sort.SliceStable(s, func(i,j int) bool { return s[i].k < s[j].k }) }
