// vybe-test: go/slices_sort_equal_extended/slices_equal_struct_elements
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
type P struct { X int }
func main() { _ = slices.Equal([]P{{1}}, []P{{1}}) }
