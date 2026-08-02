// vybe-test: go/slices_sort_equal_extended/slices_replace_string_elements
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []string{"a", "c"}
_ = slices.Replace(s, 1, 2, "b") }
