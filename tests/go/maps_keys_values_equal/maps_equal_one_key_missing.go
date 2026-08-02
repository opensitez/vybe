// vybe-test: go/maps_keys_values_equal/maps_equal_one_key_missing
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Equal(map[int]int{1: 1, 2: 2}, map[int]int{1: 1, 2: 9}) }
