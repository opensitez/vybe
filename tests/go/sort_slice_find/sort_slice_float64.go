// vybe-test: go/sort_slice_find/sort_slice_float64
// origin: languages/go/tests/go/test_sort_slice_find.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { s := []float64{1.2,0.1}
sort.Slice(s, func(i,j int) bool { return s[i] < s[j] }) }
