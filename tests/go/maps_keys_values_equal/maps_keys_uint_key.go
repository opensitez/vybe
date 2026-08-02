// vybe-test: go/maps_keys_values_equal/maps_keys_uint_key
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Keys(map[uint]int{1: 1, 2: 2}) }
