// vybe-test: go/maps_keys_values_equal/maps_equal_pointer_values
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { x := 1
a := map[int]*int{1: &x}
b := map[int]*int{1: &x}
_ = maps.Equal(a, b) }
