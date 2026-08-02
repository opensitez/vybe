// vybe-test: go/maps_keys_values_equal/maps_keys_nested_map_value
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Keys(map[string]map[int]int{"a": {1: 1}}) }
