// vybe-test: go/sort_package/sort_search_float64s
// origin: languages/go/tests/go/test_sort_package.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { _ = sort.SearchFloat64s([]float64{1.0,2.0}, 1.5) }
