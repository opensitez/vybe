// vybe-test: go/slices_sort_equal_extended/slices_index_bool_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { _ = slices.Index([]bool{true, false}, false) }
