// vybe-test: go/maps_keys_values_equal/maps_values_array_value
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Values(map[int][2]int{1: {1, 2}}) }
