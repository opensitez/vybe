// vybe-test: go/slices_maps_stdlib/slices_insert_string_slice
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []string{"a"}
_ = slices.Insert(s, 1, "b", "c") }
