// vybe-test: go/maps_keys_values_equal/maps_equal_int8_keys
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Equal(map[int8]int{1: 1}, map[int8]int{1: 1}) }
