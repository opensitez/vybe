// vybe-test: go/maps_keys_values_equal/maps_equal_func_uint_values
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.EqualFunc(map[int]uint{1: 5}, map[int]uint{1: 5}, func(x, y uint) bool { return x == y }) }
