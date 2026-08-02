// vybe-test: go/slices_sort_equal_extended/slices_index_func_pointer_elements
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []*int{new(int)}
_ = slices.IndexFunc(s, func(p *int) bool { return p != nil }) }
