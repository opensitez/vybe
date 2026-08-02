// vybe-test: go/sort_stable_search/sort_stable_by_struct_field
// origin: languages/go/tests/go/test_sort_stable_search.rs
// vybe-test-mode: compile

package main
import "sort"
type node struct { val, pri int }
func main() { s := []node{{2, 1}, {1, 2}}
sort.SliceStable(s, func(i, j int) bool { return s[i].pri < s[j].pri }) }
