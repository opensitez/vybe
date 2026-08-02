// vybe-test: go/sort_stable_search/sort_interface_reverse
// origin: languages/go/tests/go/test_sort_stable_search.rs
// vybe-test-mode: compile

package main
import "sort"
type ints []int
func (p ints) Len() int { return len(p) }
func (p ints) Less(i, j int) bool { return p[i] < p[j] }
func (p ints) Swap(i, j int) { p[i], p[j] = p[j], p[i] }
func main() { sort.Sort(sort.Reverse(ints{3, 1, 2})) }
