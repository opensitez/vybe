// vybe-test: go/sort_stable_search/sort_search_float64s_insert
// origin: languages/go/tests/go/test_sort_stable_search.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { _ = sort.SearchFloat64s([]float64{0.5, 1.5}, 1.0) }
