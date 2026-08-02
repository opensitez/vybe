// vybe-test: go/maps_keys_values_equal/maps_keys_float64_key
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Keys(map[float64]int{1.5: 1}) }
