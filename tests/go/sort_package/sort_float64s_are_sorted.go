// vybe-test: go/sort_package/sort_float64s_are_sorted
// origin: languages/go/tests/go/test_sort_package.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { _ = sort.Float64sAreSorted([]float64{1.0}) }
