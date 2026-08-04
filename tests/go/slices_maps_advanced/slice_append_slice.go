// vybe-test: go/slices_maps_advanced/slice_append_slice
// origin: languages/go/tests/go/test_slices_maps_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { s1 := []int{1, 2}
s2 := []int{3, 4}
s1 = append(s1, s2...) }
