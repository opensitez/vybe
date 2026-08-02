// vybe-test: go/sort_package/sort_interface_sort
// origin: languages/go/tests/go/test_sort_package.rs
// vybe-test-mode: compile

package main
import "sort"
type ints []int
func (p ints) Len() int { return len(p) }
func (p ints) Less(i,j int) bool { return p[i] < p[j] }
func (p ints) Swap(i,j int) { p[i], p[j] = p[j], p[i] }
func main() { sort.Sort(ints{2,1}) }
