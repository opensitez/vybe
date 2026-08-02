// vybe-test: go/maps_keys_values_equal/maps_values_interface_value
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Values(map[int]interface{}{1: 42, 2: "x"}) }
