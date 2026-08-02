// vybe-test: go/slices_maps_stdlib/slices_grow_byte_slice
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := make([]byte, 1, 1)
_ = slices.Grow(s, 4) }
