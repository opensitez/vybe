// vybe-test: go/slices_sort_equal_extended/slices_sort_byte_slice
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []byte{'z', 'a', 'm'}
slices.Sort(s) }
