// vybe-test: go/sort_package/sort_reverse_ints
// origin: languages/go/tests/go/test_sort_package.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { a := []int{1,2,3}
sort.Sort(sort.Reverse(sort.IntSlice(a))) }
